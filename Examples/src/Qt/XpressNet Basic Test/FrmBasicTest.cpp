//
//  FrmBasicTest.cpp
//  XpressNetQt examples: XpressNet Basic Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of the example XpressNet Basic Test (VB6: B_UserInterfaceFormConnect.frm, .NET: FrmBasicTest.vb), for
//  XpressNetQt 2.0. The layout (FrmBasicTest.ui) is converted from the .NET form; the code follows the .NET code.
//  Differences with .NET:
//  - The events of the XpressNet object are Qt signals, connected in the constructor (.NET: Handles).
//  - The LEDs are Led widgets (Led.cpp); the control array LEDF0F68 is filled in InitControlArrays.
//  - USB port: the name of the serial device, for example /dev/ttyUSB0 (.NET: a COM port number). The list contains
//    the serial ports that Qt finds (QSerialPortInfo); another name can be typed.
//  - The logfile directory is ~/TMC-Logfiles (.NET: C:\Users\Public\TMC-Logfiles).
//  - FormClosed / XpressNet.Dispose() is the destructor (delete XpressNet).
//  - LabelMasterVersion (a QLabel) has no Click event: the click is caught by eventFilter.
//  New in version 2.0 (as in .NET): the option Handshake for USB, and the events SendError, InterfaceNichtAdressiert
//  and InterfaceWiederAdressiert (shown in LabelStatus).
//

#include "FrmBasicTest.h"
#include "ui_FrmBasicTest.h"
#include "Led.h"
#include "XpressNetLogFile.h"
#include <QDir>
#include <QMouseEvent>
#include <QSerialPortInfo>

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

// The serial ports that Qt finds (QSerialPortInfo). USB devices have a vendor identifier: with usbFirst these are
// listed first, otherwise last. Without any port, defaultPort is shown
static void fillPorts(QComboBox *combo, bool usbFirst, const QString &defaultPort)
{
  for (const QSerialPortInfo &info : QSerialPortInfo::availablePorts()) {
    if (info.hasVendorIdentifier() == usbFirst) combo->addItem(info.systemLocation());
  }
  for (const QSerialPortInfo &info : QSerialPortInfo::availablePorts()) {
    if (info.hasVendorIdentifier() != usbFirst) combo->addItem(info.systemLocation());
  }
  if (combo->count() == 0) combo->setCurrentText(defaultPort);
}


FrmBasicTest::FrmBasicTest(QWidget *parent) : QWidget(parent), ui(new Ui::FrmBasicTest)
{
  ui->setupUi(this);
  for (int i = 0; i <= 10; i++) { FB[i] = 0; FS[i] = 0; }
  InitControlArrays();
  XpressNet = new XpressNetClass(this);
  // Set the directory for logfiles (on Linux the default directory of the library)
  XpressNet->Logfile->setLogDirectory(QDir(QDir::homePath()).filePath("TMC-Logfiles"));
  // The serial ports that Qt finds, the USB devices first. Another name can be typed
  fillPorts(ui->TxtUsbPort, true, "/dev/ttyUSB0");
  // Possible Baudrates for the USB interface
  ui->ComboBaudUsb->addItem("57600");
  ui->ComboBaudUsb->addItem("115200");
  // Possible lok decoder steps. The text of the .ui file (28) is kept
  QString steps = ui->ComboLokSteps->currentText();
  ui->ComboLokSteps->addItem("14");
  ui->ComboLokSteps->addItem("27");
  ui->ComboLokSteps->addItem("28");
  ui->ComboLokSteps->addItem("128");
  ui->ComboLokSteps->setCurrentText(steps);
  // The buttons
  connect(ui->cmdTcpConnect, &QPushButton::clicked, this, &FrmBasicTest::cmdTcpConnect_Click);
  connect(ui->CmdUsbConnect, &QPushButton::clicked, this, &FrmBasicTest::CmdUsbConnect_Click);
  ui->LabelMasterVersion->installEventFilter(this);
  connect(ui->CmdInterfaceInfo, &QPushButton::clicked, this, &FrmBasicTest::CmdInterfaceInfo_Click);
  connect(ui->CmdLanInfo, &QPushButton::clicked, this, &FrmBasicTest::CmdLanInfo_Click);
  connect(ui->cmdRecht, &QPushButton::clicked, this, &FrmBasicTest::cmdRecht_Click);
  connect(ui->cmdAf, &QPushButton::clicked, this, &FrmBasicTest::cmdAf_Click);
  connect(ui->cmdSwitchFeedback, &QPushButton::clicked, this, &FrmBasicTest::cmdSwitchFeedback_Click);
  connect(ui->CmdRsFeedback, &QPushButton::clicked, this, &FrmBasicTest::CmdRsFeedback_Click);
  connect(ui->CmdAllesAn, &QPushButton::clicked, this, &FrmBasicTest::CmdAllesAn_Click);
  connect(ui->CmdAllesAus, &QPushButton::clicked, this, &FrmBasicTest::CmdAllesAus_Click);
  connect(ui->CmdAlleLoksAnhalten, &QPushButton::clicked, this, &FrmBasicTest::CmdAlleLoksAnhalten_Click);
  connect(ui->CmdEineLokAnhalten, &QPushButton::clicked, this, &FrmBasicTest::CmdEineLokAnhalten_Click);
  connect(ui->CmdLokEnter, &QPushButton::clicked, this, &FrmBasicTest::CmdLokEnter_Click);
  connect(ui->CmdLokAusStack, &QPushButton::clicked, this, &FrmBasicTest::CmdLokAusStack_Click);
  connect(ui->CmdLokForward, &QPushButton::clicked, this, &FrmBasicTest::CmdLokForward_Click);
  connect(ui->CmdLokReverse, &QPushButton::clicked, this, &FrmBasicTest::CmdLokReverse_Click);
  connect(ui->CmdFunctionOn, &QPushButton::clicked, this, &FrmBasicTest::CmdFunctionOn_Click);
  connect(ui->CmdFunctionOff, &QPushButton::clicked, this, &FrmBasicTest::CmdFunctionOff_Click);
  connect(ui->CmdGetF0_F12, &QPushButton::clicked, this, &FrmBasicTest::CmdGetF0_F12_Click);
  connect(ui->CmdGetF13_28, &QPushButton::clicked, this, &FrmBasicTest::CmdGetF13_28_Click);
  connect(ui->CmdGetF29_68, &QPushButton::clicked, this, &FrmBasicTest::CmdGetF29_68_Click);
  // The events of the XpressNet object
  connect(XpressNet, &XpressNetClass::InterfaceConnected, this, &FrmBasicTest::XpressNet_InterfaceConnected);
  connect(XpressNet, &XpressNetClass::TcpClosedByRemoteHost, this, &FrmBasicTest::XpressNet_TcpClosedByRemoteHost);
  connect(XpressNet, &XpressNetClass::TcpError, this, &FrmBasicTest::XpressNet_TcpError);
  connect(XpressNet, &XpressNetClass::Feedback, this, &FrmBasicTest::XpressNet_Feedback);
  connect(XpressNet, &XpressNetClass::SwitchFeedback, this, &FrmBasicTest::XpressNet_SwitchFeedback);
  connect(XpressNet, &XpressNetClass::AllesAn, this, &FrmBasicTest::XpressNet_AllesAn);
  connect(XpressNet, &XpressNetClass::NotAus, this, &FrmBasicTest::XpressNet_NotAus);
  connect(XpressNet, &XpressNetClass::NotHalt, this, &FrmBasicTest::XpressNet_NotHalt);
  connect(XpressNet, &XpressNetClass::ProgrammierMode, this, &FrmBasicTest::XpressNet_ProgrammierMode);
  connect(XpressNet, &XpressNetClass::ZentraleVersion, this, &FrmBasicTest::XpressNet_ZentraleVersion);
  connect(XpressNet, &XpressNetClass::Fehlermeldung, this, &FrmBasicTest::XpressNet_Fehlermeldung);
  connect(XpressNet, &XpressNetClass::XpressNetAddress, this, &FrmBasicTest::XpressNet_XpressNetAddress);
  connect(XpressNet, &XpressNetClass::XpressNetVersion, this, &FrmBasicTest::XpressNet_XpressNetVersion);
  connect(XpressNet, &XpressNetClass::InterfaceVersion, this, &FrmBasicTest::XpressNet_InterfaceVersion);
  connect(XpressNet, &XpressNetClass::InterfaceFreieVerbindungen, this, &FrmBasicTest::XpressNet_InterfaceFreieVerbindungen);
  connect(XpressNet, &XpressNetClass::NormaleLokinfo, this, &FrmBasicTest::XpressNet_NormaleLokinfo);
  connect(XpressNet, &XpressNetClass::FunctionsZustandF13F20, this, &FrmBasicTest::XpressNet_FunctionsZustandF13F20);
  connect(XpressNet, &XpressNetClass::FunctionsZustandF29F68, this, &FrmBasicTest::XpressNet_FunctionsZustandF29F68);
  connect(XpressNet, &XpressNetClass::SendError, this, &FrmBasicTest::XpressNet_SendError);
  connect(XpressNet, &XpressNetClass::InterfaceNichtAdressiert, this, &FrmBasicTest::XpressNet_InterfaceNichtAdressiert);
  connect(XpressNet, &XpressNetClass::InterfaceWiederAdressiert, this, &FrmBasicTest::XpressNet_InterfaceWiederAdressiert);
}

// .NET: FormClosed / XpressNet.Dispose()
FrmBasicTest::~FrmBasicTest()
{
  delete XpressNet;
  delete ui;
}

// A click on LabelMasterVersion
bool FrmBasicTest::eventFilter(QObject *watched, QEvent *event)
{
  if ((watched == ui->LabelMasterVersion) && (event->type() == QEvent::MouseButtonRelease)) {
    LabelMasterVersion_Click();
    return true;
  }
  return QWidget::eventFilter(watched, event);
}


// ==================================================================================================================
void FrmBasicTest::cmdTcpConnect_Click()
{
  ui->LabelStatus->setText("TCP connecting .....");
  XpressNet->CmdConnectTcp(ui->txtIP->text(), ui->txtTcpPort->text());
}

void FrmBasicTest::CmdUsbConnect_Click()
{
  ui->LabelStatus->setText("USB connecting .....");
  QString Settings;
  Settings = ui->ComboBaudUsb->currentText() + ",N,8,1";
  XpressNet->CmdConnectUsb(ui->TxtUsbPort->currentText(), Settings, ui->CheckHandshake->isChecked());
}

void FrmBasicTest::LabelMasterVersion_Click()
{
  XpressNet->CmdSoftwareversionZentraleAnfordern();
}

void FrmBasicTest::CmdInterfaceInfo_Click()
{
  XpressNet->CmdInterfaceVersion();
}

void FrmBasicTest::CmdLanInfo_Click()
{
  XpressNet->CmdXpressNetAddress();
  XpressNet->CmdXpressNetVersion();
  XpressNet->CmdInterfaceFreieVerbindungen();
}

// ==================================================================================================================
void FrmBasicTest::cmdRecht_Click()
{
  if (!IsNumeric(ui->txtWissel->text())) return;
  XpressNet->CmdAccessoryAuto(toInt(ui->txtWissel->text()), 1);
}

void FrmBasicTest::cmdAf_Click()
{
  if (!IsNumeric(ui->txtWissel->text())) return;
  XpressNet->CmdAccessoryAuto(toInt(ui->txtWissel->text()), 0);
}

void FrmBasicTest::cmdSwitchFeedback_Click()
{
  if (IsNumeric(ui->txtWissel->text())) {
    XpressNet->CmdSwitchInfoRequest(toInt(ui->txtWissel->text()));
  }
}

void FrmBasicTest::CmdRsFeedback_Click()
{
  if (IsNumeric(ui->TxtRsBusAddress->text())) {
    XpressNet->CmdFeedbackRequest(toInt(ui->TxtRsBusAddress->text()), 0);   // Nibble 0
    XpressNet->CmdFeedbackRequest(toInt(ui->TxtRsBusAddress->text()), 1);   // Nibble 1
  }
}

void FrmBasicTest::CmdAllesAn_Click()
{
  XpressNet->CmdAllesAn();
}

void FrmBasicTest::CmdAllesAus_Click()
{
  XpressNet->CmdAllesAus();
}

void FrmBasicTest::CmdAlleLoksAnhalten_Click()
{
  XpressNet->CmdAlleLoksAnhalten();
}

void FrmBasicTest::CmdEineLokAnhalten_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  XpressNet->CmdEineLokAnhalten(toInt(ui->txtLokAddress->text()));
}

void FrmBasicTest::CmdLokEnter_Click()
{
  // A new lok address is entered. Get the lok info
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  ClearLedsForFunctions();
  XpressNet->CmdLokInfoAnfordern(toInt(ui->txtLokAddress->text()));
}

void FrmBasicTest::CmdLokAusStack_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  XpressNet->CmdLokAusStack(toInt(ui->txtLokAddress->text()));
}

void FrmBasicTest::CmdLokForward_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokSpeed->text())) return;
  XpressNet->CmdFahrbefehl(toInt(ui->txtLokAddress->text()), toInt(ui->TxtLokSpeed->text()), toInt(ui->ComboLokSteps->currentText()), true);
}

void FrmBasicTest::CmdLokReverse_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokSpeed->text())) return;
  XpressNet->CmdFahrbefehl(toInt(ui->txtLokAddress->text()), toInt(ui->TxtLokSpeed->text()), toInt(ui->ComboLokSteps->currentText()), false);
}

void FrmBasicTest::CmdFunctionOn_Click()
{
  handleFunctionBefehl(true);
}

void FrmBasicTest::CmdFunctionOff_Click()
{
  handleFunctionBefehl(false);
}

void FrmBasicTest::CmdGetF0_F12_Click()
{
  XpressNet->CmdLokInfoAnfordern(toInt(ui->txtLokAddress->text()));
}

void FrmBasicTest::CmdGetF13_28_Click()
{
  XpressNet->CmdFunktionsZustandF13F28Anfordern(toInt(ui->txtLokAddress->text()));
}

void FrmBasicTest::CmdGetF29_68_Click()
{
  XpressNet->CmdFunktionsZustandF29F68Anfordern(toInt(ui->txtLokAddress->text()));
}


// ==================================================================================================================
// EVENTS
void FrmBasicTest::XpressNet_InterfaceConnected()
{
  ui->ledInterfaceStatus->setFillColor(Qt::green);
  ui->LabelStatus->setText("Connection established");
}

void FrmBasicTest::XpressNet_TcpClosedByRemoteHost()
{
  ui->ledInterfaceStatus->setFillColor(Qt::red);
  ui->LabelStatus->setText("TCP connection closed by remote host");
}

void FrmBasicTest::XpressNet_TcpError(QString ErrorDescription)
{
  ui->ledInterfaceStatus->setFillColor(Qt::red);
  ui->LabelStatus->setText(ErrorDescription);
}

void FrmBasicTest::XpressNet_Feedback(int address, quint8 nibble, quint8 data)
{
  ui->TxBoxtRsAddress->setPlainText(QString::number(address) + "\n" + ui->TxBoxtRsAddress->toPlainText());
  ui->TxtBoxRsData->setPlainText(printBits(data) + "\n" + ui->TxtBoxRsData->toPlainText());
  ui->TxtBoxRsNibble->setPlainText(QString::number(nibble) + "\n" + ui->TxtBoxRsNibble->toPlainText());
}

void FrmBasicTest::XpressNet_SwitchFeedback(int address, quint8 ttBits, quint8 iBit, quint8 data)
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

void FrmBasicTest::XpressNet_AllesAn()
{
  ui->ledNotAus->setFillColor(Qt::green);
  ui->ledNotHalt->setFillColor(Qt::green);
  ui->ledProgrammiermode->setFillColor(Qt::green);
}

void FrmBasicTest::XpressNet_NotAus()
{
  ui->ledNotAus->setFillColor(Qt::red);
}

void FrmBasicTest::XpressNet_NotHalt()
{
  ui->ledNotHalt->setFillColor(Qt::red);
}

void FrmBasicTest::XpressNet_ProgrammierMode()
{
  ui->ledProgrammiermode->setFillColor(Qt::red);
}

void FrmBasicTest::XpressNet_ZentraleVersion(QString version)
{
  ui->txtMasterVersion->setText(version);
}

void FrmBasicTest::XpressNet_Fehlermeldung(quint8 Value)
{
  ui->LabelStatus->setText("Fehlermeldung: " + QString::number(Value) + " (V3.6: 3.1.14)");
}

void FrmBasicTest::XpressNet_XpressNetAddress(quint8 address)
{
  ui->txtXAddress->setText(QString::number(address));
}

void FrmBasicTest::XpressNet_XpressNetVersion(QString version)
{
  ui->txtXVersion->setText(version);
}

void FrmBasicTest::XpressNet_InterfaceVersion(QString version, QString code)
{
  ui->txtInterfaceVersion->setText(version);
  ui->txtInterfaceCode->setText(code);
}

void FrmBasicTest::XpressNet_InterfaceFreieVerbindungen(quint8 number)
{
  ui->txtTcpFree->setText(QString::number(number));
}

void FrmBasicTest::XpressNet_NormaleLokinfo(bool Besetzt, int Speedsteps, int Speed, bool Forward, quint8 F0_F4, quint8 F5_F8, quint8 F9_F12)
{
  Q_UNUSED(Besetzt);
  Q_UNUSED(Forward);
  // Set the combobox and speed, based on the received input
  ui->ComboLokSteps->setCurrentText(QString::number(Speedsteps));
  ui->TxtLokSpeed->setText(QString::number(Speed));
  // Store the values of F0..F12, to allow changing a single Function later
  FB[1] = F0_F4;                                          // Gruppe 1
  FB[2] = F5_F8;                                          // Gruppe 2
  FB[3] = F9_F12;                                         // Gruppe 3
  // Show the LEDs for F0 .. F12
  ShowFbGroup(F0_F4, 1);                                  // Led FB bits for Gruppe 1
  ShowFbGroup(F5_F8, 2);                                  // Led FB bits for Gruppe 2
  ShowFbGroup(F9_F12, 3);                                 // Led FB bits for Gruppe 3
}

void FrmBasicTest::XpressNet_FunctionsZustandF13F20(quint8 F13_F20, quint8 F21_F28)
{
  FB[4] = F13_F20;                                        // Store Gruppe 4
  FB[5] = F21_F28;                                        // Store Gruppe 5
  ShowFbGroup(F13_F20, 4);                                // Led FB bits for Gruppe 4
  ShowFbGroup(F21_F28, 5);                                // Led FB bits for Gruppe 5
}

void FrmBasicTest::XpressNet_FunctionsZustandF29F68(quint8 F29_F36, quint8 F37_F44, quint8 F45_F52, quint8 F53_F60, quint8 F61_F68)
{
  FB[6] = F29_F36;                                        // Store Gruppe 6
  FB[7] = F37_F44;                                        // Store Gruppe 7
  FB[8] = F45_F52;                                        // Store Gruppe 8
  FB[9] = F53_F60;                                        // Store Gruppe 9
  FB[10] = F61_F68;                                       // Store Gruppe 10
  ShowFbGroup(F29_F36, 6);                                // Led FB bits for Gruppe 6
  ShowFbGroup(F37_F44, 7);                                // Led FB bits for Gruppe 7
  ShowFbGroup(F45_F52, 8);                                // Led FB bits for Gruppe 8
  ShowFbGroup(F53_F60, 9);                                // Led FB bits for Gruppe 9
  ShowFbGroup(F61_F68, 10);                               // Led FB bits for Gruppe 10
}

// Version 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum
void FrmBasicTest::XpressNet_SendError(quint8 Reason, QString Description)
{
  ui->LabelStatus->setText(sendErrorName(Reason) + ": " + Description);
}

// Version 2.0: the command station does not address the interface: sending is paused
void FrmBasicTest::XpressNet_InterfaceNichtAdressiert()
{
  ui->LabelStatus->setText("Interface nicht adressiert");
}

// Version 2.0: the command station addresses the interface again
void FrmBasicTest::XpressNet_InterfaceWiederAdressiert()
{
  ui->LabelStatus->setText("Interface wieder adressiert");
}


// ==================================================================================================================
// Support routines below
// ------------------------------------------------------------------------------------------------------------------
// The VB6 control array LEDF0F68: the Led widgets LEDF0F68_0 .. LEDF0F68_68 of the form
void FrmBasicTest::InitControlArrays()
{
  for (int i = 0; i <= 68; i++) LEDF0F68[i] = findChild<Led *>("LEDF0F68_" + QString::number(i));
}

// ------------------------------------------------------------------------------------------------------------------
// Lok Functions
void FrmBasicTest::handleFunctionBefehl(bool turnOn)
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

void FrmBasicTest::ClearLedsForFunctions()
{
  for (int i = 0; i <= 68; i++) {
    LEDF0F68[i]->setFillColor(Qt::white);
  }
}

QColor FrmBasicTest::FbColor(bool Value)
{
  if (Value) return Qt::red;
  else return Qt::green;
}

// ------------------------------------------------------------------------------------------------------------------
// Functions to show the Feedbacks on this form
void FrmBasicTest::ShowFbGroup(quint8 data, quint8 group)
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
bool FrmBasicTest::bitValue(quint8 data, quint8 index)
{
  return ((data >> index) & 1) == 1;
}

quint8 FrmBasicTest::bitSet(quint8 data, quint8 index)
{
  return quint8(data | (1 << index));
}

quint8 FrmBasicTest::bitClear(quint8 data, quint8 index)
{
  return quint8(data & (255 - (1 << index)));            // mask the data
}

quint8 FrmBasicTest::bitUpdate(quint8 data, quint8 index, bool bitIsSet)
{
  if (bitIsSet) return bitSet(data, index);
  else return bitClear(data, index);
}

// ------------------------------------------------------------------------------------------------------------------
// Transforms numbers form 0..15 (4 bits) in a bitstring
QString FrmBasicTest::printBits(quint8 data)
{
  QString outString;
  for (int i = 0; i <= 3; i++) {
    outString = QString::number(data % 2) + outString;
    data = data / 2;
  }
  return outString;
}
