//
//  DecoderScanner.cpp
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See DecoderScanner.h for a description of the three steps of the scan.
//  Only one PoM read is outstanding at any moment, since all decoders answer via the same RS-Bus address (128).
//

#include "DecoderScanner.h"
#include "XpressNet.h"
#include <algorithm>

#define CONNECT_TIMEOUT_TIME 5.0                // Seconds to wait for the connections with the Lenz interfaces
#define RSBUS_TIMEOUT_TIME 20.0                 // Seconds to wait for the RS-Bus data of one command station
#define POM_TIMEOUT_TIME 1.0                    // Seconds to wait for the RS-Bus answer on a PoM read (as Decoder-POM)
#define MAX_ATTEMPTS 2                          // Number of PoM reads per CV (per command station for CV7)
#define QUIET_TIME 0.5                          // Seconds to ignore all answers after a timeout or a bad value
#define FEC_GUARD_TIME 0.3                      // Seconds to ignore repeated answers (CV7 and RSFEC > 0)
#define MAX_REREADS 2                           // Number of extra reads of a CV with a value that is not possible
#define ACCESSORY_POM_OFFSET 6999               // PoM address = 6999 + RS-Bus address (CV10 = decoder address + 1)
#define FEEDBACK_POM_OFFSET 6000                // PoM address = 6000 + RS-Bus address (as Programmer GBM-POM)
#define RS_BUS_FEEDBACK_ADDRESS 128             // The RS-Bus address for the PoM answers
#define RSBUS_PART 32                           // RS-Bus addresses per Get RS-Bus data request (see nextRsStation)
#define PROGRESS_RSBUS 10.0                     // Part (%) of the progress bar for step 2; step 3 takes the rest

// The names of the CVs, as in CvValues.h of the AP_DCC_Decoder_Core library
#define myAddrL        1
#define version        7
#define myAddrH        9
#define myRSAddr      10
#define CmdStation    19
#define RSFEC         20
#define SkipUnEven    21
#define DecType       27
#define SwitchCount   35
#define SwitchType    36


// ****************************************************************************************************************
// *********************************************** SCANNED DECODER ************************************************
// ****************************************************************************************************************
ScannedDecoder::ScannedDecoder()
{
  for (int i = 0; i <= MAX_CV_NUMBER; i++) cvStore.append(-1);
}

int ScannedDecoder::pomAddress() const
{
  if (givenPomAddress > 0) return givenPomAddress;
  if (rsBusType == ACCESSORY_DECODER) return rsAddress + ACCESSORY_POM_OFFSET;
  return rsAddress + FEEDBACK_POM_OFFSET;
}

void ScannedDecoder::setCv(int cvNumber, int value)
{
  if ((cvNumber > 0) && (cvNumber <= MAX_CV_NUMBER)) cvStore[cvNumber] = value;
  // A decoder with a given PoM address gets its RS-Bus address from CV10
  if ((givenPomAddress > 0) && (cvNumber == myRSAddr)) rsAddress = value;
}

int ScannedDecoder::getCv(int cvNumber) const
{
  if ((cvNumber > 0) && (cvNumber <= MAX_CV_NUMBER)) return cvStore[cvNumber];
  return -1;
}

QString ScannedDecoder::cvText(int cvNumber) const
{
  int value = getCv(cvNumber);
  if (value < 0) return "?";
  return QString::number(value);
}

// The number of switches of an accessory decoder. Decoders with version < 20 have 4 switches. From version 20,
// as AVR-Servo-3-6.ino: SwitchType (CV36) 1 (DKW) = 2 switches, 2 (double crossover) = 4 switches, otherwise
// SwitchCount (CV35). SwitchCount 0, 255 or above the board maximum means: the board maximum.
// A GBM with relays board (DecType 50) has 4 relays (NUMBER_OF_DEVICES in global.h of the GBM firmware).
// The lift decoder (DecType 24) listens to 3 decoder addresses = 12 switch addresses (Lift_Main.ino)
int ScannedDecoder::numberOfSwitches() const
{
  if (rsBusType != ACCESSORY_DECODER) return (cvStore[DecType] == 50) ? 4 : -1;
  int cv7 = cvStore[version];
  if (cv7 < 0) return -1;
  if (cvStore[DecType] == 24) return 12;
  if (cv7 < 20) return 4;
  int boardMaximum;
  switch (cvStore[DecType]) {
    case 20:  boardMaximum = 2; break;          // Servo-2
    case 21:  boardMaximum = 3; break;          // Servo-3
    case 22:  boardMaximum = 6; break;          // Servo-6
    case 33:  boardMaximum = 16; break;         // Relays decoder for 16 relais
    default:  boardMaximum = 4; break;
  }
  if (cvStore[SwitchType] == 1) return 2;
  if (cvStore[SwitchType] == 2) return 4;
  int count = cvStore[SwitchCount];
  if ((count <= 0) || (count == 255) || (count > boardMaximum)) return boardMaximum;
  return count;
}

// An RS-Bus address holds 4 switches (2 bits each). With SkipUnEven each switch has its own nibble (2 per address).
// The lift decoder uses two RS-Bus addresses: myRSAddr and myRSAddr + 1 (feedback.cpp of Lift_Main)
int ScannedDecoder::rsAddressCount() const
{
  if (liftIr) return 2;                         // RS_AddresLow and RS_AddresHigh (mySettings.h of Lift_IR)
  if (rsBusType != ACCESSORY_DECODER) return 1;
  if (cvStore[DecType] == 24) return 2;
  int switches = numberOfSwitches();
  if (switches <= 0) return 1;
  if (cvStore[SkipUnEven] == 1) return (switches + 1) / 2;
  return (switches + 3) / 4;
}

QString ScannedDecoder::rsAddressText() const
{
  if (rsAddress == 0) return "?";               // A decoder with a given PoM address, CV10 not received
  int count = rsAddressCount();
  if (count <= 1) return QString::number(rsAddress);
  return QString("%1-%2").arg(rsAddress).arg(rsAddress + count - 1);
}

// The switch addresses of the decoder (first address as getSwitchAddress of Programmer Decoder-POM).
// With SkipUnEven each switch uses two switch addresses (of which only the even one is used).
// A GBM listens to the accessory address in CV1 and CV9 for its relays (main.c and dcc_decode.c of the GBM
// firmware). It is shown for a GBM with relays board only. If bit 7 of CV9 is set, the decoder has not been
// initialised (main.c of the GBM and switch decoder firmware): "-"
QString ScannedDecoder::accessoryAddressText() const
{
  int cv1 = cvStore[myAddrL];
  int cv9 = cvStore[myAddrH];
  int cv7 = cvStore[version];
  int cv27 = cvStore[DecType];
  int decoderAddress;
  if (rsBusType != ACCESSORY_DECODER) {
    if (cv27 < 0) return "?";
    if (cv27 != 50) return "";                  // 50: GBM with relays board
  }
  if ((cv1 < 0) || (cv9 < 0) || (cv7 < 0)) return "?";
  if (cv9 & 0x80) return "-";
  if (rsBusType == ACCESSORY_DECODER) {
    // Decoders with version >= 16: decoderAddress = CV9 << 6 + CV1 - 1 (as Programmer Decoder-POM).
    // The lift decoder has version 12, but uses storedAddress() of AP_DCC_Decoder_Core, which also subtracts 1
    decoderAddress = ((cv9 & 0x7F) << 6) | cv1;
    if ((cv7 >= 16) || (cv27 == 24)) decoderAddress = decoderAddress - 1;
  }
  else decoderAddress = (((cv9 & 0x07) << 6) | cv1) - 1;
  if (decoderAddress < 0) return "-";
  int first = decoderAddress * 4 + 1;
  int switches = numberOfSwitches();
  if ((cvStore[SkipUnEven] == 1) && (cv27 != 24)) switches = switches * 2;
  return QString("%1-%2").arg(first).arg(first + switches - 1);
}

// CmdStation (CV19), as CvValues.h of the AP_DCC_Decoder_Core library
QString ScannedDecoder::cmdStationText() const
{
  switch (cvStore[CmdStation]) {
    case -1: return "?";
    case 0:  return "Roco";
    case 1:  return "Lenz";
    case 2:  return "OpenDCC Z1";
  }
  return QString::number(cvStore[CmdStation]);
}

// The names of DecType (CV27), as in Programmer Decoder-POM (DCCDecoder.m) and CvValues.h of AP_DCC_Decoder_Core
QString ScannedDecoder::typeName(int value)
{
  QString type;
  QString subType;
  switch (value) {
    case 16:  type = "Switch decoder"; break;
    case 17:  type = "Switch decoder"; subType = "with Emergency board"; break;
    case 20:  type = "Servo-2 Decoder"; break;
    case 21:  type = "Servo-3 Decoder"; break;
    case 22:  type = "Servo-6 Decoder"; break;
    case 24:  type = "Lift decoder - Level"; break;
    case 32:  type = "Relays decoder"; subType = "for 4 relais"; break;
    case 33:  type = "Relays decoder"; subType = "for 16 relais"; break;
    case 34:  type = "Relays decoder"; subType = "for a configurable number of relais"; break;
    case 48:  type = "Track Occupancy decoder"; break;
    case 49:  type = "Track Occupancy decoder"; subType = "with reverser board"; break;
    case 50:  type = "Track Occupancy decoder"; subType = "with relays board"; break;
    case 52:  type = "Track Occupancy decoder"; subType = "with speed measurement"; break;
    case 64:  type = "Function Decoder"; break;
    case 128: type = "Safety decoder"; subType = "with Watchdog"; break;
    case 193: type = "TMC 24 Channel IO Decoder"; break;
    case 194: type = "TMC 16 Channel Switch Decoder"; break;
    default:  return QString();
  }
  if (subType.isEmpty()) return type;
  return type + " " + subType;
}

QString ScannedDecoder::typeText() const
{
  if (liftIr) return "Lift decoder - IR";
  int value = cvStore[DecType];
  if (value < 0) return "?";
  QString name = typeName(value);
  if (name.isEmpty()) return QString("Unknown (CV27 = %1)").arg(value);
  return name;
}


// ****************************************************************************************************************
// *********************************************** DECODER SCANNER ************************************************
// ****************************************************************************************************************
DecoderScanner::DecoderScanner(QObject *parent) : QObject(parent)
{
  timer.setSingleShot(true);
  connect(&timer, &QTimer::timeout, this, &DecoderScanner::timerFired);
}

DecoderScanner::~DecoderScanner()
{
  timer.stop();
  for (Station *station : stations) delete station->XpressNet;   // Closes the connection
  qDeleteAll(stations);
  qDeleteAll(retired);
  qDeleteAll(owned);
}

ScannedDecoder *DecoderScanner::newDecoder()
{
  ScannedDecoder *decoder = new ScannedDecoder();
  owned.append(decoder);
  return decoder;
}


// ************************************************************************************************************
// ************************************************ START AND STOP ********************************************
// ************************************************************************************************************
void DecoderScanner::start(const QStringList &ipAddresses, const QList<int> &cvNumbers, int pomAddress, int irAddress)
{
  if (phase != IDLE) return;
  // The objects of the previous scan
  qDeleteAll(retired);
  retired.clear();
  found.clear();
  decoderList.clear();
  liftRows.clear();
  qDeleteAll(owned);
  owned.clear();
  liftPomAddress = pomAddress;
  liftIrAddress = irAddress;
  userCvs.clear();
  for (int cv : cvNumbers) {
    if ((cv >= 1) && (cv <= MAX_CV_NUMBER) && !userCvs.contains(cv)) userCvs.append(cv);
  }
  problems.clear();
  noAnswer.clear();
  lastSuccessfulStation = 0;
  // Step 1: create a XpressNet object for every IP address, and open the connection. The events of the XpressNet
  // objects are connected to lambda functions, that add the station
  for (int i = 0; i < ipAddresses.count(); i++) {
    QString ipAddress = ipAddresses[i].trimmed();
    if (ipAddress.isEmpty()) continue;
    Station *station = new Station;
    station->number = i + 1;
    station->ipAddress = ipAddress;
    XpressNetClass *xpressNet = new XpressNetClass(this);
    station->XpressNet = xpressNet;
    connect(xpressNet, &XpressNetClass::InterfaceConnected, this, [this, station]() {
      stationConnected(station, true, QString()); });
    connect(xpressNet, &XpressNetClass::TcpClosedByRemoteHost, this, [this, station]() {
      stationConnected(station, false, "TCP connection closed by remote host"); });
    connect(xpressNet, &XpressNetClass::TcpError, this, [this, station](QString Description) {
      stationConnected(station, false, Description); });
    // Feedback decoder (TT=10): address = 1..128
    connect(xpressNet, &XpressNetClass::Feedback, this, [this, station](int address, quint8 Nibble, quint8 data) {
      stationAnswer(station, address, Nibble, data, 2); });
    // Switch decoder (TT=00, 01 or 11): switchAddress = ((address * 4) + (Nibble * 2) + 1), with address = 0..127.
    // The conversion back to RS-Bus address and nibble is as in the RS-Bus Monitor
    connect(xpressNet, &XpressNetClass::SwitchFeedback, this, [this, station](int switchAddress, quint8 ttBits, quint8, quint8 data) {
      int rsAddress = (switchAddress - 1) / 4 + 1;
      int nibble = ((switchAddress - 1) / 2) % 2;
      stationAnswer(station, rsAddress, nibble, data, ttBits); });
    connect(xpressNet, &XpressNetClass::SendError, this, [this, station](quint8 Reason, QString Description) {
      stationSendError(station, QString("SendError %1: %2").arg(Reason).arg(Description)); });
    stations.append(station);
  }
  if (stations.isEmpty()) {
    emit scannerFinished("No IP address in the preferences");
    return;
  }
  phase = CONNECTING;
  emit scannerStatus(QString("Connecting to %1 command station(s) .....").arg(stations.count()), 0);
  for (Station *station : stations) station->XpressNet->CmdConnectTcp(station->ipAddress, "5550");
  // A connection may already have failed in CmdConnectTcp (for example: an invalid IP address)
  if (phase == CONNECTING) startTimeout(CONNECT_TIMEOUT_TIME, CONNECT_TIMEOUT);
}

void DecoderScanner::stop()
{
  if (phase == IDLE) return;
  finish("Scan stopped");
}

// Sort order of the table: on the station that delivered the PoM messages, then on PoM address
static bool compareDecoders(const ScannedDecoder *d1, const ScannedDecoder *d2)
{
  if (d1->pomStation != d2->pomStation) return d1->pomStation < d2->pomStation;
  return d1->pomAddress() < d2->pomAddress();
}

// The end of the scan. The XpressNet objects (and thus the connections) are deleted after the current event has
// been handled, since this method may be called from within an event of one of these objects. The Station
// objects are deleted at the next scan, for the same reason
void DecoderScanner::finish(const QString &text)
{
  timer.stop();
  phase = IDLE;
  waitingForAnswer = false;
  for (Station *station : stations) {
    disconnect(station->XpressNet, nullptr, this, nullptr);     // Later events are ignored
    station->XpressNet->deleteLater();
    station->XpressNet = nullptr;
    station->connected = false;
  }
  retired.append(stations);
  stations.clear();
  candidates.clear();
  std::stable_sort(decoderList.begin(), decoderList.end(), compareDecoders);
  decoderList.append(liftRows);                 // The lift decoder at the end of the table
  QString finalText = text;
  if (!problems.isEmpty()) finalText = text + ". " + problems.join("; ");
  emit scannerFinished(finalText);
}


// ************************************************************************************************************
// **************************************************** TIMER *************************************************
// ************************************************************************************************************
void DecoderScanner::startTimeout(double seconds, timerAction_t action)
{
  timerAction = action;
  timer.start(qRound(seconds * 1000));
}

void DecoderScanner::timerFired()
{
  switch (timerAction) {
    case CONNECT_TIMEOUT: connectTimeout(); break;
    case RSBUS_TIMEOUT:   rsBusTimeout(); break;
    case POM_TIMEOUT:     pomTimeout(); break;
    // The next step after the quiet time or the FEC guard time. In the meantime all answers are ignored
    case RESEND:          if (phase == POM_READ) sendPomRead(); break;
    case NEXT_CV:         if (phase == POM_READ) nextCv(); break;
    case NEXT_DECODER:    if (phase == POM_READ) nextDecoder(); break;
  }
}


// ************************************************************************************************************
// ******************************************** STEP 1: CONNECT ***********************************************
// ************************************************************************************************************
void DecoderScanner::stationConnected(Station *station, bool ok, const QString &text)
{
  if (phase == CONNECTING) {
    if (station->connectDone) return;
    station->connectDone = true;
    station->connected = ok;
    if (!ok) problems.append(QString("Command station %1 (%2): %3").arg(station->number).arg(station->ipAddress, text));
    for (Station *s : stations) if (!s->connectDone) return;
    startRsBusData();
    return;
  }
  // The connection is closed during the scan. Step 3 continues with the other stations (via the timeouts)
  if (!ok && station->connected) {
    station->connected = false;
    problems.append(QString("Command station %1 (%2): %3").arg(station->number).arg(station->ipAddress, text));
    if ((phase == RSBUS_DATA) && (station == stations[rsStationIndex])) nextRsStation();
  }
}

void DecoderScanner::connectTimeout()
{
  if (phase != CONNECTING) return;
  for (Station *station : stations) {
    if (!station->connectDone) {
      station->connectDone = true;
      station->connected = false;
      problems.append(QString("Command station %1 (%2): no connection").arg(station->number).arg(station->ipAddress));
    }
  }
  startRsBusData();
}


// ************************************************************************************************************
// ****************************************** STEP 2: RS-BUS DATA *********************************************
// ************************************************************************************************************
void DecoderScanner::startRsBusData()
{
  timer.stop();
  numberConnected = 0;
  for (Station *station : stations) if (station->connected) numberConnected++;
  if (numberConnected == 0) {
    finish("No connection with any command station");
    return;
  }
  phase = RSBUS_DATA;
  rsStationIndex = -1;
  rsStationsDone = 0;
  nextRsStation();
}

// As Get RS-Bus data of the RS-Bus Monitor, but in four parts of 32 RS-Bus addresses (two nibbles each). The FIFO of
// the library holds 128 messages, but after the connection is established the library itself also sends messages
// (software version and status of the command station). The next part is requested after both answers for the last
// RS-Bus address of the current part; the station is ready after both answers for RS-Bus address 128
void DecoderScanner::nextRsStation()
{
  timer.stop();
  if (rsStationIndex >= 0) rsStationsDone++;
  rsStationIndex++;
  while ((rsStationIndex < stations.count()) && !stations[rsStationIndex]->connected) rsStationIndex++;
  if (rsStationIndex >= stations.count()) {
    startPom();
    return;
  }
  Station *station = stations[rsStationIndex];
  partLastAddress = RSBUS_PART;
  answerCount = 0;
  showRsBusProgress(0);
  requestFeedbacks(station, 1, partLastAddress);
  startTimeout(RSBUS_TIMEOUT_TIME, RSBUS_TIMEOUT);
}

void DecoderScanner::requestFeedbacks(Station *station, int firstAddress, int lastAddress)
{
  for (int address = firstAddress; address <= lastAddress; address++) {
    station->XpressNet->CmdFeedbackRequest(address, 0);
    station->XpressNet->CmdFeedbackRequest(address, 1);
  }
}

void DecoderScanner::showRsBusProgress(int rsAddress)
{
  Station *station = stations[rsStationIndex];
  QString text = QString("Command station %1 (%2): requesting RS-Bus data .....").arg(station->number).arg(station->ipAddress);
  double progress = PROGRESS_RSBUS * (rsStationsDone + rsAddress / 128.0) / numberConnected;
  emit scannerStatus(text, progress);
}

void DecoderScanner::rsBusTimeout()
{
  Station *station = stations[rsStationIndex];
  problems.append(QString("Command station %1 (%2): RS-Bus data incomplete").arg(station->number).arg(station->ipAddress));
  nextRsStation();
}

// An answer during step 2: TT=01 (accessory decoder) and TT=10 (GBM) are stored. RS-Bus address 128 is skipped,
// since it holds the PoM answers
void DecoderScanner::rsBusAnswer(Station *station, int rsAddress, int ttBits)
{
  if (station != stations[rsStationIndex]) return;
  if ((rsAddress < RS_BUS_FEEDBACK_ADDRESS) && ((ttBits == 1) || (ttBits == 2))) {
    bool known = false;
    for (ScannedDecoder *d : found) if ((d->rsStation == station->number) && (d->rsAddress == rsAddress)) known = true;
    if (!known) {
      ScannedDecoder *decoder = newDecoder();
      decoder->rsStation = station->number;
      decoder->rsAddress = rsAddress;
      decoder->rsBusType = (ttBits == 1) ? ACCESSORY_DECODER : FEEDBACK_DECODER;
      found.append(decoder);
    }
  }
  showRsBusProgress(rsAddress);
  if (rsAddress != partLastAddress) return;
  answerCount++;
  if (answerCount < 2) return;
  if (partLastAddress >= 128) {                 // The last part has been answered
    nextRsStation();
    return;
  }
  answerCount = 0;                              // The next part
  partLastAddress = partLastAddress + RSBUS_PART;
  requestFeedbacks(station, partLastAddress - RSBUS_PART + 1, partLastAddress);
}


// ************************************************************************************************************
// ********************************************** STEP 3: POM *************************************************
// ************************************************************************************************************
// Sort order: first the accessory decoders, then the GBMs; per type on RS-Bus address
static bool compareFound(const ScannedDecoder *d1, const ScannedDecoder *d2)
{
  if (d1->rsBusType != d2->rsBusType) return d1->rsBusType < d2->rsBusType;
  if (d1->rsAddress != d2->rsAddress) return d1->rsAddress < d2->rsAddress;
  return d1->rsStation < d2->rsStation;
}

// A value that is not possible, according to the ranges in CvValues.h of the AP_DCC_Decoder_Core library.
// CV35 and CV36 differ between accessory decoders and GBMs, and are only checked for accessory decoders
static bool isPlausible(int cv, int value, rsBusType_t type)
{
  switch (cv) {
    case version:     return (value >= 8);
    case myAddrH:     return ((value & 0x7F) <= 7);         // Bit 7: not initialised
    case myRSAddr:    return (value <= 128);
    case CmdStation:  return (value <= 2);
    case RSFEC:       return (value <= 2);
    case DecType:     return !ScannedDecoder::typeName(value).isEmpty();
    case SkipUnEven:  return (value <= 1);
    case SwitchType:  return (type != ACCESSORY_DECODER) || (value <= 2);
  }
  return true;
}

void DecoderScanner::startPom()
{
  timer.stop();
  phase = POM_READ;
  std::stable_sort(found.begin(), found.end(), compareFound);
  // The lift decoder is read first, so its RS-Bus addresses are known (and skipped) during the rest of the scan.
  // Its rows are added at the end of the table, after the scan. Its RS-Bus address comes from CV10, and its answers
  // are accepted from any command station (rsStation 0) until the first answer has been received
  if (liftPomAddress > 0) {
    ScannedDecoder *decoder = newDecoder();
    decoder->rsBusType = ACCESSORY_DECODER;
    decoder->rsStation = 0;
    decoder->rsAddress = 0;
    decoder->givenPomAddress = liftPomAddress;
    found.prepend(decoder);
  }
  decoderIndex = -1;
  nextDecoder();
}

void DecoderScanner::nextDecoder()
{
  decoderIndex++;
  while ((decoderIndex < found.count()) && found[decoderIndex]->skip) decoderIndex++;
  if (decoderIndex >= found.count()) {
    if (!noAnswer.isEmpty()) problems.prepend("No PoM answer for " + noAnswer.join(", "));
    finish(QString("Scan ready: %1 decoder(s) answered via PoM").arg(decoderList.count()));
    return;
  }
  // The stations to try for CV7: the last successful one first, then the others
  candidates.clear();
  for (Station *station : stations)
    if (station->connected && (station->number == lastSuccessfulStation)) candidates.append(station);
  for (Station *station : stations)
    if (station->connected && (station->number != lastSuccessfulStation)) candidates.append(station);
  if (candidates.isEmpty()) {
    finish("No connection with any command station");
    return;
  }
  // First CV7 only. The other CVs are added once CV7 has been received (see buildCvList)
  cvList.clear();
  cvList.append(version);
  candidateIndex = 0;
  cvIndex = 0;
  attempt = 0;
  rereads = 0;
  sendPomRead();
}

// The CVs needed for the table first (RSFEC second, for the FEC guard time), then the CVs chosen by the user
void DecoderScanner::buildCvList(ScannedDecoder *decoder)
{
  cvList.append(RSFEC);
  cvList.append(DecType);
  cvList.append(SkipUnEven);
  if (decoder->rsBusType == ACCESSORY_DECODER) {
    if (decoder->getCv(version) >= 20) {
      cvList.append(SwitchCount);
      cvList.append(SwitchType);
    }
  }
  cvList.append(myAddrL);
  cvList.append(myAddrH);
  if (decoder->givenPomAddress > 0) cvList.append(myRSAddr);
  for (int cv : userCvs) if (!cvList.contains(cv)) cvList.append(cv);
}

Station *DecoderScanner::stationWithNumber(int number)
{
  for (Station *station : stations) if (station->number == number) return station;
  return nullptr;
}

void DecoderScanner::sendPomRead()
{
  ScannedDecoder *decoder = found[decoderIndex];
  Station *station;
  if (cvIndex == 0) station = candidates[candidateIndex];
  else station = stationWithNumber(decoder->pomStation);
  int cv = cvList[cvIndex];
  QString text = QString("Decoder %1 of %2 (RS-Bus address %3, PoM address %4): CV%5 via command station %6")
                   .arg(decoderIndex + 1).arg(found.count()).arg(decoder->rsAddress).arg(decoder->pomAddress()).arg(cv).arg(station->number);
  double progress = PROGRESS_RSBUS + (100.0 - PROGRESS_RSBUS) * (decoderIndex + (double)cvIndex / cvList.count()) / found.count();
  emit scannerStatus(text, progress);
  haveLowNibble = false;
  waitingForAnswer = true;
  station->XpressNet->CmdPomRead(decoder->pomAddress(), cv);
  startTimeout(POM_TIMEOUT_TIME, POM_TIMEOUT);
}

// An answer during step 3. The decoder sends two RS-Bus messages: first the low nibble, then the high nibble
void DecoderScanner::pomAnswer(Station *station, int rsAddress, int nibble, int data)
{
  if (!waitingForAnswer) return;
  ScannedDecoder *decoder = found[decoderIndex];
  if ((decoder->rsStation != 0) && (station->number != decoder->rsStation)) return;
  bool ownAddress = (decoder->rsBusType == FEEDBACK_DECODER) && (rsAddress == decoder->rsAddress);   // Old GBM firmware
  if ((rsAddress != RS_BUS_FEEDBACK_ADDRESS) && !ownAddress) return;
  if (nibble == 0) {
    lowNibble = data & 0x0F;
    haveLowNibble = true;
    return;
  }
  if (!haveLowNibble) return;                   // The first nibble was lost: wait for the timeout
  timer.stop();
  waitingForAnswer = false;
  int value = ((data & 0x0F) << 4) + lowNibble;
  int cv = cvList[cvIndex];
  // A value that is not possible is read again, after the quiet time
  if (!isPlausible(cv, value, decoder->rsBusType) && (rereads < MAX_REREADS)) {
    rereads++;
    startTimeout(QUIET_TIME, RESEND);
    return;
  }
  decoder->setCv(cv, value);
  if (decoder->rsStation == 0) decoder->rsStation = station->number;   // A decoder with a given PoM address
  if (cvIndex == 0) {
    Station *pomStation = candidates[candidateIndex];
    decoder->pomStation = pomStation->number;
    decoder->pomIpAddress = pomStation->ipAddress;
    lastSuccessfulStation = pomStation->number;
    buildCvList(decoder);
    if (decoder->givenPomAddress > 0) liftRows.append(decoder);
    else {
      decoderList.append(decoder);
      emit scannerFoundDecoder(decoder);
    }
  }
  else emit scannerUpdatedDecoder(decoder);
  // The FEC guard time after CV7 (RSFEC is not known yet), and if the decoder repeats its answers (RSFEC > 0)
  if ((cvIndex == 0) || (decoder->getCv(RSFEC) > 0)) startTimeout(FEC_GUARD_TIME, NEXT_CV);
  else nextCv();
}

void DecoderScanner::nextCv()
{
  cvIndex++;
  attempt = 0;
  rereads = 0;
  if (cvIndex < cvList.count()) {
    sendPomRead();
    return;
  }
  // All CVs have been read. The RS-Bus address(es) of this decoder are not scanned (again). For a decoder with a
  // given PoM address this includes its first RS-Bus address. The TT bits are not checked, since the RSbus library
  // sends "feedback decoder" messages by default (RSbus.h), for example for the lift decoder
  ScannedDecoder *decoder = found[decoderIndex];
  if (decoder->rsBusType == ACCESSORY_DECODER) {
    int last = decoder->rsAddress + decoder->rsAddressCount() - 1;
    for (ScannedDecoder *other : found) {
      if ((other != decoder) && (other->rsStation == decoder->rsStation) &&
          (other->rsAddress >= decoder->rsAddress) && (other->rsAddress <= last)) other->skip = true;
    }
  }
  if (decoder->givenPomAddress > 0) addLiftIr(decoder);
  nextDecoder();
}

// The IR sensors of the lift have their own board (Lift_IR), without DCC and thus without PoM. Their RS-Bus
// addresses are given in the preferences. If they are on the RS-Bus of the lift decoder, they get their own row,
// with the PoM address and CVs of the lift decoder
void DecoderScanner::addLiftIr(ScannedDecoder *lift)
{
  if (liftIrAddress <= 0) return;
  bool onRsBus = false;
  for (ScannedDecoder *other : found) {
    if ((other->rsStation == lift->rsStation) && (other->rsAddress >= liftIrAddress) && (other->rsAddress <= liftIrAddress + 1)) {
      other->skip = true;
      onRsBus = true;
    }
  }
  if (!onRsBus) {
    problems.append(QString("IR sensors of the lift not found on RS-Bus address %1").arg(liftIrAddress));
    return;
  }
  ScannedDecoder *ir = newDecoder();
  ir->liftIr = true;
  ir->liftLevel = lift;
  ir->rsStation = lift->rsStation;
  ir->rsAddress = liftIrAddress;
  ir->pomStation = lift->pomStation;
  ir->pomIpAddress = lift->pomIpAddress;
  liftRows.insert(liftRows.indexOf(lift), ir);  // The lower RS-Bus addresses first
}

void DecoderScanner::pomTimeout()
{
  waitingForAnswer = false;
  attempt++;
  if (attempt < MAX_ATTEMPTS) {
    startTimeout(QUIET_TIME, RESEND);
    return;
  }
  attempt = 0;
  rereads = 0;
  if (cvIndex == 0) {
    // No answer on CV7 via this station: try the next one. Without any answer the decoder is omitted
    candidateIndex++;
    if (candidateIndex < candidates.count()) startTimeout(QUIET_TIME, RESEND);
    else {
      ScannedDecoder *decoder = found[decoderIndex];
      if (decoder->givenPomAddress > 0) noAnswer.append(QString("PoM %1").arg(decoder->givenPomAddress));
      else {
        Station *station = stationWithNumber(decoder->rsStation);
        noAnswer.append(QString("RS-Bus %1 (%2)").arg(decoder->rsAddress).arg(station ? station->ipAddress : QString()));
      }
      startTimeout(QUIET_TIME, NEXT_DECODER);
    }
    return;
  }
  // No answer for this CV: the value stays unknown ("?")
  emit scannerUpdatedDecoder(found[decoderIndex]);
  startTimeout(QUIET_TIME, NEXT_CV);
}


// ************************************************************************************************************
// ******************************************** EVENTS OF THE STATIONS ****************************************
// ************************************************************************************************************
void DecoderScanner::stationAnswer(Station *station, int rsAddress, int nibble, int data, int ttBits)
{
  if (phase == RSBUS_DATA) rsBusAnswer(station, rsAddress, ttBits);
  else if (phase == POM_READ) pomAnswer(station, rsAddress, nibble, data);
}

// In step 2 the library has emptied its queues: the RS-Bus data of this station is incomplete.
// In step 3 the timeout handles the missing answer
void DecoderScanner::stationSendError(Station *station, const QString &text)
{
  if ((phase == RSBUS_DATA) && (station == stations[rsStationIndex])) {
    problems.append(QString("Command station %1 (%2): %3").arg(station->number).arg(station->ipAddress, text));
    nextRsStation();
  }
}
