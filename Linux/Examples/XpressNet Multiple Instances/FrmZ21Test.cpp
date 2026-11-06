//
//  FrmZ21Test.cpp
//  XpressNetQt examples: Z21 Basic Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of the example Z21 Basic Test (VB6: B_UserInterfaceFormZ21.frm, .NET: FrmZ21Test.vb), for
//  XpressNetQt 2.0. The layout (FrmZ21Test.ui) is converted from the .NET form; the code follows the .NET code.
//  Differences with .NET:
//  - The events of the XpressNet object are Qt signals, connected in the constructor (.NET: Handles).
//  - The LEDs are Led widgets (Led.cpp); the control array LEDF0F68 is filled in InitControlArrays.
//  - The logfile directory is ~/TMC-Logfiles (.NET: C:\Users\Public\TMC-Logfiles).
//  - The Timer component TimerMessageCounter is a QTimer.
//  - FormClosed / XpressNet.Dispose() is the destructor (delete XpressNet).
//  New in version 2.0 (as in .NET): the status line LabelStatus for the event SendError.
//  (The events InterfaceNichtAdressiert and InterfaceWiederAdressiert are not used by the Z21.)
//

#include "FrmZ21Test.h"
#include "ui_FrmZ21Test.h"
#include "Led.h"
#include "XpressNetLogFile.h"
#include <QDir>

// VB: IsNumeric. Is the text a number?
static bool IsNumeric(const QString &text)
{
  bool ok = false;
  text.trimmed().toDouble(&ok);
  return ok;
}

// VB: the conversion of a text to a number (for example txtWissel.Text as parameter of CmdAccessoryAuto)
static int toInt(const QString &text)
{
  return qRound(text.trimmed().toDouble());
}

// .NET: CType(Reason, SendErrorReasonEnum).ToString()
static QString sendErrorName(quint8 Reason)
{
  switch (Reason) {
    case WiederholungErfolglos: return "WiederholungErfolglos";
    case BefehlNichtVorhanden: return "BefehlNichtVorhanden";
    case Timeout: return "Timeout";
    case FifoVoll: return "FifoVoll";
    case KeineVerbindung: return "KeineVerbindung";
    case FehlerInDenBefehlsparametern: return "FehlerInDenBefehlsparametern";
    case ZentraleAdressiertDasInterfaceNichtMehr: return "ZentraleAdressiertDasInterfaceNichtMehr";
    default: return QString::number(Reason);
  }
}


FrmZ21Test::FrmZ21Test(QWidget *parent) : QWidget(parent), ui(new Ui::FrmZ21Test)
{
  ui->setupUi(this);
  for (int i = 0; i <= 10; i++) { FB[i] = 0; FS[i] = 0; }
  InitControlArrays();
  XpressNet = new XpressNetClass(this);                   // We only use the XpressNet part of Z21
  // Set the directory for logfiles (on Linux the default directory of the library)
  XpressNet->Logfile->setLogDirectory(QDir(QDir::homePath()).filePath("TMC-Logfiles"));
  // Possible lok decoder steps (as in the XpressNet Basic test). The text of the .ui file is kept
  QString steps = ui->ComboLokSteps->currentText();
  ui->ComboLokSteps->addItem("14");
  ui->ComboLokSteps->addItem("27");
  ui->ComboLokSteps->addItem("28");
  ui->ComboLokSteps->addItem("128");
  ui->ComboLokSteps->setCurrentText(steps);
  // The Timer component of the .NET form
  TimerMessageCounter.setInterval(1000);
  connect(&TimerMessageCounter, &QTimer::timeout, this, &FrmZ21Test::TimerMessageCounter_Timer);
  TimerMessageCounter.start();
  // The buttons
  connect(ui->cmdUdpSet, &QPushButton::clicked, this, &FrmZ21Test::cmdUdpSet_Click);
  connect(ui->CmdAllesAn, &QPushButton::clicked, this, &FrmZ21Test::CmdAllesAn_Click);
  connect(ui->CmdAllesAus, &QPushButton::clicked, this, &FrmZ21Test::CmdAllesAus_Click);
  connect(ui->CmdAlleLoksAnhalten, &QPushButton::clicked, this, &FrmZ21Test::CmdAlleLoksAnhalten_Click);
  connect(ui->CmdEineLokAnhalten, &QPushButton::clicked, this, &FrmZ21Test::CmdEineLokAnhalten_Click);
  connect(ui->CmdLokEnter, &QPushButton::clicked, this, &FrmZ21Test::CmdLokEnter_Click);
  connect(ui->CmdLokAusStack, &QPushButton::clicked, this, &FrmZ21Test::CmdLokAusStack_Click);
  connect(ui->CmdLokForward, &QPushButton::clicked, this, &FrmZ21Test::CmdLokForward_Click);
  connect(ui->CmdLokReverse, &QPushButton::clicked, this, &FrmZ21Test::CmdLokReverse_Click);
  connect(ui->CmdFunctionOn, &QPushButton::clicked, this, &FrmZ21Test::CmdFunctionOn_Click);
  connect(ui->CmdFunctionOff, &QPushButton::clicked, this, &FrmZ21Test::CmdFunctionOff_Click);
  connect(ui->CmdGetLokInfo, &QPushButton::clicked, this, &FrmZ21Test::CmdGetLokInfo_Click);
  connect(ui->cmdRecht, &QPushButton::clicked, this, &FrmZ21Test::cmdRecht_Click);
  connect(ui->cmdAf, &QPushButton::clicked, this, &FrmZ21Test::cmdAf_Click);
  connect(ui->cmdSwitchFeedback, &QPushButton::clicked, this, &FrmZ21Test::cmdSwitchFeedback_Click);
  connect(ui->CmdExtAccSet, &QPushButton::clicked, this, &FrmZ21Test::CmdExtAccSet_Click);
  // The events of the XpressNet object
  connect(XpressNet, &XpressNetClass::SwitchFeedback, this, &FrmZ21Test::XpressNet_SwitchFeedback);
  connect(XpressNet, &XpressNetClass::AllesAn, this, &FrmZ21Test::XpressNet_AllesAn);
  connect(XpressNet, &XpressNetClass::NotAus, this, &FrmZ21Test::XpressNet_NotAus);
  connect(XpressNet, &XpressNetClass::NotHalt, this, &FrmZ21Test::XpressNet_NotHalt);
  connect(XpressNet, &XpressNetClass::ProgrammierMode, this, &FrmZ21Test::XpressNet_ProgrammierMode);
  connect(XpressNet, &XpressNetClass::ZentraleVersion, this, &FrmZ21Test::XpressNet_ZentraleVersion);
  connect(XpressNet, &XpressNetClass::z21FirmwareVersion, this, &FrmZ21Test::XpressNet_Z21FirmwareVersion);
  connect(XpressNet, &XpressNetClass::Z21Lokinfo, this, &FrmZ21Test::XpressNet_Z21Lokinfo);
  connect(XpressNet, &XpressNetClass::SendError, this, &FrmZ21Test::XpressNet_SendError);
}

// .NET: FormClosed / XpressNet.Dispose()
FrmZ21Test::~FrmZ21Test()
{
  TimerMessageCounter.stop();
  delete XpressNet;
  delete ui;
}

void FrmZ21Test::TimerMessageCounter_Timer()
{
  ui->txtNmbPckts->setText(QString::number(XpressNet->NumberOfExpressNetMessagesReceived + XpressNet->NumberOfExpressNetMessagesTransmitted));
}


// ==================================================================================================================
void FrmZ21Test::cmdUdpSet_Click()
{
  XpressNet->CmdInitUdp(ui->txtIP->text(), ui->txtUdpPort->text());
}

void FrmZ21Test::CmdAllesAn_Click()
{
  XpressNet->CmdAllesAn();
}

void FrmZ21Test::CmdAllesAus_Click()
{
  XpressNet->CmdAllesAus();
}

void FrmZ21Test::CmdAlleLoksAnhalten_Click()
{
  XpressNet->CmdAlleLoksAnhalten();
}

void FrmZ21Test::CmdEineLokAnhalten_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  XpressNet->CmdEineLokAnhalten(toInt(ui->txtLokAddress->text()));
}

void FrmZ21Test::CmdLokEnter_Click()
{
  // A new lok address is entered. Get the lok info
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  ClearLedsForFunctions();
  XpressNet->CmdLokInfoAnfordern(toInt(ui->txtLokAddress->text()));
}

void FrmZ21Test::CmdLokAusStack_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  XpressNet->CmdLokAusStack(toInt(ui->txtLokAddress->text()));
}

void FrmZ21Test::CmdLokForward_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokSpeed->text())) return;
  XpressNet->CmdFahrbefehl(toInt(ui->txtLokAddress->text()), toInt(ui->TxtLokSpeed->text()), toInt(ui->ComboLokSteps->currentText()), true);
}

void FrmZ21Test::CmdLokReverse_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokSpeed->text())) return;
  XpressNet->CmdFahrbefehl(toInt(ui->txtLokAddress->text()), toInt(ui->TxtLokSpeed->text()), toInt(ui->ComboLokSteps->currentText()), false);
}

void FrmZ21Test::CmdFunctionOn_Click()
{
  handleFunctionBefehl(true);
}

void FrmZ21Test::CmdFunctionOff_Click()
{
  handleFunctionBefehl(false);
}

void FrmZ21Test::CmdGetLokInfo_Click()
{
  XpressNet->CmdLokInfoAnfordern(toInt(ui->txtLokAddress->text()));
}


// ==================================================================================================================
void FrmZ21Test::cmdRecht_Click()
{
  if (!IsNumeric(ui->txtWissel->text())) return;
  if (ui->CheckAutoDeactivate->isChecked()) {
    XpressNet->CmdAccessoryAuto(toInt(ui->txtWissel->text()), 1);
  }
  else {
    XpressNet->CmdAccessory(toInt(ui->txtWissel->text()), 1, ui->CheckActivate->isChecked() ? 1 : 0);
    if (ui->CheckActivate->isChecked()) {
      ui->CheckActivate->setChecked(false);
    }
    else {
      ui->CheckActivate->setChecked(true);
    }
  }
}

void FrmZ21Test::cmdAf_Click()
{
  if (!IsNumeric(ui->txtWissel->text())) return;
  if (ui->CheckAutoDeactivate->isChecked()) {
    XpressNet->CmdAccessoryAuto(toInt(ui->txtWissel->text()), 0);
  }
  else {
    XpressNet->CmdAccessory(toInt(ui->txtWissel->text()), 0, ui->CheckActivate->isChecked() ? 1 : 0);
    if (ui->CheckActivate->isChecked()) {
      ui->CheckActivate->setChecked(false);
    }
    else {
      ui->CheckActivate->setChecked(true);
    }
  }
}

void FrmZ21Test::cmdSwitchFeedback_Click()
{
  if (IsNumeric(ui->txtWissel->text())) {
    XpressNet->CmdSwitchInfoRequest(toInt(ui->txtWissel->text()));
  }
}

void FrmZ21Test::CmdExtAccSet_Click()
{
  if (IsNumeric(ui->TxtExtAccSwitch->text())) {
    // XpressNet->CmdOpenDCCSetExtAccessory(toInt(ui->TxtExtAccSwitch->text()), toInt(ui->TxtExtAccValue->text()));
    XpressNet->CmdZ21SetExtAccessory(toInt(ui->TxtExtAccSwitch->text()), toInt(ui->TxtExtAccValue->text()));
  }
}


// ==================================================================================================================
// EVENTS
void FrmZ21Test::XpressNet_SwitchFeedback(int address, quint8 ttBits, quint8 iBit, quint8 data)
{
  QString ttText;
  if (ttBits == 0) {
    ttText = "C";                                         // Feedback comes from Command Station
  }
  else {
    ttText = "D";                                         // Feedback comes from Decoder
  }
  ui->TxtBoxSwitchAddress->setPlainText(QString::number(address) + "/" + QString::number(address + 1) + "\n" + ui->TxtBoxSwitchAddress->toPlainText());
  ui->TxtBoxSwitchData->setPlainText(printBits(data) + "\n" + ui->TxtBoxSwitchData->toPlainText());
  ui->TxtBoxSwitchTtBits->setPlainText(ttText + "\n" + ui->TxtBoxSwitchTtBits->toPlainText());
  ui->TxtBoxSwitchIBit->setPlainText(QString::number(iBit) + "\n" + ui->TxtBoxSwitchIBit->toPlainText());
}

void FrmZ21Test::XpressNet_AllesAn()
{
  ui->ledNotAus->setFillColor(Qt::green);
  ui->ledNotHalt->setFillColor(Qt::green);
  ui->ledProgrammiermode->setFillColor(Qt::green);
}

void FrmZ21Test::XpressNet_NotAus()
{
  ui->ledNotAus->setFillColor(Qt::red);
}

void FrmZ21Test::XpressNet_NotHalt()
{
  ui->ledNotHalt->setFillColor(Qt::red);
}

void FrmZ21Test::XpressNet_ProgrammierMode()
{
  ui->ledProgrammiermode->setFillColor(Qt::red);
}

void FrmZ21Test::XpressNet_ZentraleVersion(QString version)
{
  ui->txtMasterVersion->setText(version);
}

void FrmZ21Test::XpressNet_Z21FirmwareVersion(QString version)
{
  ui->txtXZ21Firmware->setText(version);
}

void FrmZ21Test::XpressNet_Z21Lokinfo(int LokAddress, bool Besetzt, int Speedsteps, int Speed, bool Forward, quint8 F0_F4, quint8 F5_F8, quint8 F9_F12, quint8 F13_F20, quint8 F21_F28)
{
  Q_UNUSED(LokAddress);
  Q_UNUSED(Besetzt);
  Q_UNUSED(Forward);
  // Set the combobox and speed, based on the received input
  ui->ComboLokSteps->setCurrentText(QString::number(Speedsteps));
  ui->TxtLokSpeed->setText(QString::number(Speed));
  // Store the values of F0..F12, to allow changing a single Function later
  FB[1] = F0_F4;                                          // Store Gruppe 1
  FB[2] = F5_F8;                                          // Store Gruppe 2
  FB[3] = F9_F12;                                         // Store Gruppe 3
  FB[4] = F13_F20;                                        // Store Gruppe 4
  FB[5] = F21_F28;                                        // Store Gruppe 5
  // Show the LEDs for F0 .. F28
  ShowFbGroup(F0_F4, 1);                                  // Led FB bits for Gruppe 1
  ShowFbGroup(F5_F8, 2);                                  // Led FB bits for Gruppe 2
  ShowFbGroup(F9_F12, 3);                                 // Led FB bits for Gruppe 3
  ShowFbGroup(F13_F20, 4);                                // Led FB bits for Gruppe 4
  ShowFbGroup(F21_F28, 5);                                // Led FB bits for Gruppe 5
}

// Version 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum
void FrmZ21Test::XpressNet_SendError(quint8 Reason, QString Description)
{
  ui->LabelStatus->setText(sendErrorName(Reason) + ": " + Description);
}


// ==================================================================================================================
// Support routines below
// ------------------------------------------------------------------------------------------------------------------
// The VB6 control array LEDF0F68: the Led widgets LEDF0F68_0 .. LEDF0F68_28 of the form
void FrmZ21Test::InitControlArrays()
{
  for (int i = 0; i <= 28; i++) LEDF0F68[i] = findChild<Led *>("LEDF0F68_" + QString::number(i));
}

// ------------------------------------------------------------------------------------------------------------------
// Lok Functions
void FrmZ21Test::handleFunctionBefehl(bool turnOn)
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokFNumber->text())) return;
  int FbNumber;
  FbNumber = toInt(ui->TxtLokFNumber->text());
  if (FbNumber < 0) return;
  if (FbNumber > 68) return;
  quint8 group = 0;
  quint8 bitposition = 0;
  // F0 needs to be consideren as speciaal case
  if (FbNumber == 0) { group = 1; bitposition = 4; }
  // The other functions are coded in a consistent way
  if ((FbNumber >= 1) && (FbNumber <= 4)) { group = 1; bitposition = FbNumber - 1; }
  if ((FbNumber >= 5) && (FbNumber <= 8)) { group = 2; bitposition = FbNumber - 5; }
  if ((FbNumber >= 9) && (FbNumber <= 12)) { group = 3; bitposition = FbNumber - 9; }
  if ((FbNumber >= 13) && (FbNumber <= 20)) { group = 4; bitposition = FbNumber - 13; }
  if ((FbNumber >= 21) && (FbNumber <= 28)) { group = 5; bitposition = FbNumber - 21; }
  if ((FbNumber >= 29) && (FbNumber <= 36)) { group = 6; bitposition = FbNumber - 29; }
  if ((FbNumber >= 37) && (FbNumber <= 44)) { group = 7; bitposition = FbNumber - 37; }
  if ((FbNumber >= 45) && (FbNumber <= 52)) { group = 8; bitposition = FbNumber - 45; }
  if ((FbNumber >= 53) && (FbNumber <= 60)) { group = 9; bitposition = FbNumber - 53; }
  if ((FbNumber >= 61) && (FbNumber <= 68)) { group = 10; bitposition = FbNumber - 61; }
  FB[group] = bitUpdate(FB[group], bitposition, turnOn);
  XpressNet->CmdLokFunktionBefehl(toInt(ui->txtLokAddress->text()), group, FB[group]);
}

void FrmZ21Test::ClearLedsForFunctions()
{
  for (int i = 0; i <= 28; i++) {
    LEDF0F68[i]->setFillColor(Qt::white);
  }
}

QColor FrmZ21Test::FbColor(bool Value)
{
  if (Value) return Qt::red;
  else return Qt::green;
}

// ------------------------------------------------------------------------------------------------------------------
// Functions to show the Feedbacks on this form
void FrmZ21Test::ShowFbGroup(quint8 data, quint8 group)
{
  int bitsInGroup;                                        // Groups 1..3 contain 4 FB bits, the others contain 8 FB bits
  int offset;                                             // Offset in the Array for a certain group
  if (group < 4) {
    bitsInGroup = 4;
    offset = group * bitsInGroup - 4;
  }
  else {
    bitsInGroup = 8;
    offset = group * bitsInGroup - 20;
  }
  for (int i = 1; i <= bitsInGroup; i++) {
    LEDF0F68[i + offset]->setFillColor(FbColor(bitValue(data, quint8(i - 1))));
  }
  // FB0 in Group 1 is a special case
  if (group == 1) LEDF0F68[0]->setFillColor(FbColor(bitValue(data, 4)));
}

// ------------------------------------------------------------------------------------------------------------------
// Some bit manipulation functions
bool FrmZ21Test::bitValue(quint8 data, quint8 index)
{
  return ((data >> index) & 1) == 1;
}

quint8 FrmZ21Test::bitSet(quint8 data, quint8 index)
{
  return quint8(data | (1 << index));
}

quint8 FrmZ21Test::bitClear(quint8 data, quint8 index)
{
  return quint8(data & (255 - (1 << index)));            // mask the data
}

quint8 FrmZ21Test::bitUpdate(quint8 data, quint8 index, bool bitIsSet)
{
  if (bitIsSet) return bitSet(data, index);
  else return bitClear(data, index);
}

// ------------------------------------------------------------------------------------------------------------------
// Transforms numbers form 0..15 (4 bits) in a bitstring
QString FrmZ21Test::printBits(quint8 data)
{
  QString outString;
  for (int i = 0; i <= 3; i++) {
    outString = QString::number(data % 2) + outString;
    data = data / 2;
  }
  return outString;
}
