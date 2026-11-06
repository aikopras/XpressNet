//
//  FrmExtendedTest.cpp
//  XpressNetQt examples: XpressNet Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of the example XpressNet Test (VB6: B_UserInterfaceFormConnect.frm, .NET: FrmExtendedTest.vb), for
//  XpressNetQt 2.0. The layout (FrmExtendedTest.ui) is converted from the .NET form; the code follows the .NET code.
//  Differences with .NET:
//  - The events of the XpressNet object are Qt signals, connected in the constructor (.NET: Handles).
//  - The LEDs are Led widgets (Led.cpp); the control array LEDF0F68 is filled in InitControlArrays.
//  - USB and serial port: the name of the serial device, for example /dev/ttyUSB0 or /dev/ttyS0 (.NET: a COM port
//    number). The lists contain the serial ports that Qt finds (QSerialPortInfo); another name can be typed.
//  - The Timer component TimerMessageCounter is a QTimer.
//  - The logfile directory is ~/TMC-Logfiles (.NET: C:\Users\Public\TMC-Logfiles).
//  - FormClosed / XpressNet.Dispose() is the destructor (delete XpressNet).
//  - LabelMasterVersion (a QLabel) has no Click event: the click is caught by eventFilter.
//  New in version 2.0 (as in .NET): the option Handshake for USB, and the events SendError, InterfaceNichtAdressiert
//  and InterfaceWiederAdressiert (shown in LabelStatus).
//

#include "FrmExtendedTest.h"
#include "ui_FrmExtendedTest.h"
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


FrmExtendedTest::FrmExtendedTest(QWidget *parent) : QWidget(parent), ui(new Ui::FrmExtendedTest)
{
  ui->setupUi(this);
  for (int i = 0; i <= 10; i++) { FB[i] = 0; FS[i] = 0; }
  InitControlArrays();
  XpressNet = new XpressNetClass(this);
  // Set the directory for logfiles (on Linux the default directory of the library)
  XpressNet->Logfile->setLogDirectory(QDir(QDir::homePath()).filePath("TMC-Logfiles"));
  // The serial ports that Qt finds, the USB devices first. Another name can be typed
  fillPorts(ui->TxtUsbPort, true, "/dev/ttyUSB0");
  fillPorts(ui->TxtSerialPort, false, "/dev/ttyS0");
  // Possible Baudrates for the USB interface
  ui->ComboBaudUsb->addItem("57600");
  ui->ComboBaudUsb->addItem("115200");
  // Possible Baudrates for the serial interface. The text of the .ui file (19200) is kept
  QString baud = ui->ComboBaud->currentText();
  ui->ComboBaud->addItem("9600");
  ui->ComboBaud->addItem("19200");
  ui->ComboBaud->addItem("38400");
  ui->ComboBaud->addItem("57600");
  ui->ComboBaud->addItem("115200");
  ui->ComboBaud->setCurrentText(baud);
  // Possible days for the Modellzeit
  QString day = ui->ComboDay->currentText();
  ui->ComboDay->addItem("Monday");
  ui->ComboDay->addItem("Tuesday");
  ui->ComboDay->addItem("Wednesday");
  ui->ComboDay->addItem("Thursday");
  ui->ComboDay->addItem("Friday");
  ui->ComboDay->addItem("Saterday");
  ui->ComboDay->addItem("Sunday");
  ui->ComboDay->setCurrentText(day);
  // Possible lok decoder steps. The text of the .ui file (28) is kept
  QString steps = ui->ComboLokSteps->currentText();
  ui->ComboLokSteps->addItem("14");
  ui->ComboLokSteps->addItem("27");
  ui->ComboLokSteps->addItem("28");
  ui->ComboLokSteps->addItem("128");
  ui->ComboLokSteps->setCurrentText(steps);
  // The Timer component of the .NET form
  TimerMessageCounter.setInterval(1000);
  connect(&TimerMessageCounter, &QTimer::timeout, this, &FrmExtendedTest::TimerMessageCounter_Timer);
  TimerMessageCounter.start();
  // The buttons
  connect(ui->cmdTcpConnect, &QPushButton::clicked, this, &FrmExtendedTest::cmdTcpConnect_Click);
  connect(ui->CmdUsbConnect, &QPushButton::clicked, this, &FrmExtendedTest::CmdUsbConnect_Click);
  connect(ui->CmdSerialConnect, &QPushButton::clicked, this, &FrmExtendedTest::CmdSerialConnect_Click);
  ui->LabelMasterVersion->installEventFilter(this);
  connect(ui->CmdInterfaceInfo, &QPushButton::clicked, this, &FrmExtendedTest::CmdInterfaceInfo_Click);
  connect(ui->CmdLanInfo, &QPushButton::clicked, this, &FrmExtendedTest::CmdLanInfo_Click);
  connect(ui->CmdErweiterteVersion, &QPushButton::clicked, this, &FrmExtendedTest::CmdErweiterteVersion_Click);
  connect(ui->CmdStartMan, &QPushButton::clicked, this, &FrmExtendedTest::CmdStartMan_Click);
  connect(ui->cmdStartAuto, &QPushButton::clicked, this, &FrmExtendedTest::cmdStartAuto_Click);
  connect(ui->cmdRecht, &QPushButton::clicked, this, &FrmExtendedTest::cmdRecht_Click);
  connect(ui->cmdAf, &QPushButton::clicked, this, &FrmExtendedTest::cmdAf_Click);
  connect(ui->cmdSwitchFeedback, &QPushButton::clicked, this, &FrmExtendedTest::cmdSwitchFeedback_Click);
  connect(ui->CmdRsFeedback, &QPushButton::clicked, this, &FrmExtendedTest::CmdRsFeedback_Click);
  connect(ui->CmdExtAccSet, &QPushButton::clicked, this, &FrmExtendedTest::CmdExtAccSet_Click);
  connect(ui->Cmdtest, &QPushButton::clicked, this, &FrmExtendedTest::CmdTest_Click);
  connect(ui->CmdAllesAn, &QPushButton::clicked, this, &FrmExtendedTest::CmdAllesAn_Click);
  connect(ui->CmdAllesAus, &QPushButton::clicked, this, &FrmExtendedTest::CmdAllesAus_Click);
  connect(ui->CmdAlleLoksAnhalten, &QPushButton::clicked, this, &FrmExtendedTest::CmdAlleLoksAnhalten_Click);
  connect(ui->CmdEineLokAnhalten, &QPushButton::clicked, this, &FrmExtendedTest::CmdEineLokAnhalten_Click);
  connect(ui->CmdLokEnter, &QPushButton::clicked, this, &FrmExtendedTest::CmdLokEnter_Click);
  connect(ui->CmdLokAusStack, &QPushButton::clicked, this, &FrmExtendedTest::CmdLokAusStack_Click);
  connect(ui->CmdLokForward, &QPushButton::clicked, this, &FrmExtendedTest::CmdLokForward_Click);
  connect(ui->CmdLokReverse, &QPushButton::clicked, this, &FrmExtendedTest::CmdLokReverse_Click);
  connect(ui->CmdFunctionOn, &QPushButton::clicked, this, &FrmExtendedTest::CmdFunctionOn_Click);
  connect(ui->CmdFunctionOff, &QPushButton::clicked, this, &FrmExtendedTest::CmdFunctionOff_Click);
  connect(ui->CmdGetF0_F12, &QPushButton::clicked, this, &FrmExtendedTest::CmdGetF0_F12_Click);
  connect(ui->CmdGetF13_28, &QPushButton::clicked, this, &FrmExtendedTest::CmdGetF13_28_Click);
  connect(ui->CmdGetF29_68, &QPushButton::clicked, this, &FrmExtendedTest::CmdGetF29_68_Click);
  connect(ui->CmdPomGet, &QPushButton::clicked, this, &FrmExtendedTest::CmdPomGet_Click);
  connect(ui->CmdPomSet, &QPushButton::clicked, this, &FrmExtendedTest::CmdPomSet_Click);
  connect(ui->cmdPomHolen, &QPushButton::clicked, this, &FrmExtendedTest::cmdPomHolen_Click);
  connect(ui->CmdModellzeitAnfordern, &QPushButton::clicked, this, &FrmExtendedTest::CmdModellzeitAnfordern_Click);
  connect(ui->CmdModellzeitAnhalten, &QPushButton::clicked, this, &FrmExtendedTest::CmdModellzeitAnhalten_Click);
  connect(ui->CmdModellzeitStarten, &QPushButton::clicked, this, &FrmExtendedTest::CmdModellzeitStarten_Click);
  connect(ui->CmdModellzeitStellen, &QPushButton::clicked, this, &FrmExtendedTest::CmdModellzeitStellen_Click);
  // The events of the XpressNet object
  connect(XpressNet, &XpressNetClass::InterfaceConnected, this, &FrmExtendedTest::XpressNet_InterfaceConnected);
  connect(XpressNet, &XpressNetClass::TcpClosedByRemoteHost, this, &FrmExtendedTest::XpressNet_TcpClosedByRemoteHost);
  connect(XpressNet, &XpressNetClass::TcpError, this, &FrmExtendedTest::XpressNet_TcpError);
  connect(XpressNet, &XpressNetClass::Feedback, this, &FrmExtendedTest::XpressNet_Feedback);
  connect(XpressNet, &XpressNetClass::SwitchFeedback, this, &FrmExtendedTest::XpressNet_SwitchFeedback);
  connect(XpressNet, &XpressNetClass::AllesAn, this, &FrmExtendedTest::XpressNet_AllesAn);
  connect(XpressNet, &XpressNetClass::NotAus, this, &FrmExtendedTest::XpressNet_NotAus);
  connect(XpressNet, &XpressNetClass::NotHalt, this, &FrmExtendedTest::XpressNet_NotHalt);
  connect(XpressNet, &XpressNetClass::ProgrammierMode, this, &FrmExtendedTest::XpressNet_ProgrammierMode);
  connect(XpressNet, &XpressNetClass::ProgrammierinfoCV, this, &FrmExtendedTest::XpressNet_ProgrammierinfoCV);
  connect(XpressNet, &XpressNetClass::ProgrammierinfoPageRegister, this, &FrmExtendedTest::XpressNet_ProgrammierinfoPageRegister);
  connect(XpressNet, &XpressNetClass::ZentraleVersion, this, &FrmExtendedTest::XpressNet_ZentraleVersion);
  connect(XpressNet, &XpressNetClass::ErweiterteZentraleVersion, this, &FrmExtendedTest::XpressNet_ErweiterteZentraleVersion);
  connect(XpressNet, &XpressNetClass::Fehlermeldung, this, &FrmExtendedTest::XpressNet_Fehlermeldung);
  connect(XpressNet, &XpressNetClass::XpressNetAddress, this, &FrmExtendedTest::XpressNet_XpressNetAddress);
  connect(XpressNet, &XpressNetClass::XpressNetVersion, this, &FrmExtendedTest::XpressNet_XpressNetVersion);
  connect(XpressNet, &XpressNetClass::InterfaceVersion, this, &FrmExtendedTest::XpressNet_InterfaceVersion);
  connect(XpressNet, &XpressNetClass::InterfaceFreieVerbindungen, this, &FrmExtendedTest::XpressNet_InterfaceFreieVerbindungen);
  connect(XpressNet, &XpressNetClass::Modellzeit, this, &FrmExtendedTest::XpressNet_Modellzeit);
  connect(XpressNet, &XpressNetClass::NormaleLokinfo, this, &FrmExtendedTest::XpressNet_NormaleLokinfo);
  connect(XpressNet, &XpressNetClass::FunctionsZustandF13F20, this, &FrmExtendedTest::XpressNet_FunctionsZustandF13F20);
  connect(XpressNet, &XpressNetClass::FunctionsZustandF29F68, this, &FrmExtendedTest::XpressNet_FunctionsZustandF29F68);
  connect(XpressNet, &XpressNetClass::SendError, this, &FrmExtendedTest::XpressNet_SendError);
  connect(XpressNet, &XpressNetClass::InterfaceNichtAdressiert, this, &FrmExtendedTest::XpressNet_InterfaceNichtAdressiert);
  connect(XpressNet, &XpressNetClass::InterfaceWiederAdressiert, this, &FrmExtendedTest::XpressNet_InterfaceWiederAdressiert);
}

// .NET: FormClosed / XpressNet.Dispose()
FrmExtendedTest::~FrmExtendedTest()
{
  TimerMessageCounter.stop();
  delete XpressNet;
  delete ui;
}

// A click on LabelMasterVersion
bool FrmExtendedTest::eventFilter(QObject *watched, QEvent *event)
{
  if ((watched == ui->LabelMasterVersion) && (event->type() == QEvent::MouseButtonRelease)) {
    LabelMasterVersion_Click();
    return true;
  }
  return QWidget::eventFilter(watched, event);
}


// ==================================================================================================================
void FrmExtendedTest::cmdTcpConnect_Click()
{
  ui->LabelStatus->setText("TCP connecting .....");
  XpressNet->CmdConnectTcp(ui->txtIP->text(), ui->txtTcpPort->text());
}

void FrmExtendedTest::CmdUsbConnect_Click()
{
  ui->LabelStatus->setText("USB connecting .....");
  QString Settings;
  Settings = ui->ComboBaudUsb->currentText() + ",N,8,1";
  XpressNet->CmdConnectUsb(ui->TxtUsbPort->currentText(), Settings, ui->CheckHandshake->isChecked());
}

void FrmExtendedTest::CmdSerialConnect_Click()
{
  ui->LabelStatus->setText("Serial connecting .....");
  QString Settings;
  Settings = ui->ComboBaud->currentText() + ",N,8,1";
  XpressNet->CmdConnectSerial(ui->TxtSerialPort->currentText(), Settings, ui->CheckCts->isChecked());
}

void FrmExtendedTest::LabelMasterVersion_Click()
{
  XpressNet->CmdSoftwareversionZentraleAnfordern();
}

void FrmExtendedTest::CmdInterfaceInfo_Click()
{
  XpressNet->CmdInterfaceVersion();
}

void FrmExtendedTest::CmdLanInfo_Click()
{
  XpressNet->CmdXpressNetAddress();
  XpressNet->CmdXpressNetVersion();
  XpressNet->CmdInterfaceFreieVerbindungen();
}

void FrmExtendedTest::CmdErweiterteVersion_Click()
{
  XpressNet->CmdErweiterteZentralenVersionsinformation();
}

void FrmExtendedTest::CmdStartMan_Click()
{
  XpressNet->CmdZentralenStartmodeSetzen(false);
}

void FrmExtendedTest::cmdStartAuto_Click()
{
  XpressNet->CmdZentralenStartmodeSetzen(true);
}

// ==================================================================================================================
void FrmExtendedTest::cmdRecht_Click()
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

void FrmExtendedTest::cmdAf_Click()
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

void FrmExtendedTest::cmdSwitchFeedback_Click()
{
  if (IsNumeric(ui->txtWissel->text())) {
    XpressNet->CmdSwitchInfoRequest(toInt(ui->txtWissel->text()));
  }
}

void FrmExtendedTest::CmdRsFeedback_Click()
{
  if (IsNumeric(ui->TxtRsBusAddress->text())) {
    XpressNet->CmdFeedbackRequest(toInt(ui->TxtRsBusAddress->text()), 0);   // Nibble 0
    XpressNet->CmdFeedbackRequest(toInt(ui->TxtRsBusAddress->text()), 1);   // Nibble 1
  }
}

void FrmExtendedTest::CmdExtAccSet_Click()
{
  if (IsNumeric(ui->TxtExtAccSwitch->text())) {
    // XpressNet->CmdOpenDCCSetExtAccessory(toInt(ui->TxtExtAccSwitch->text()), toInt(ui->TxtExtAccValue->text()));
    XpressNet->CmdZ21SetExtAccessory(toInt(ui->TxtExtAccSwitch->text()), toInt(ui->TxtExtAccValue->text()));
  }
}

void FrmExtendedTest::CmdTest_Click()
{
  // Test
  XpressNet->CmdAccessory(745, 0, 1);
  XpressNet->CmdAccessory(745, 0, 0);
  XpressNet->CmdAccessory(745, 1, 1);
  XpressNet->CmdAccessory(745, 1, 0);
  XpressNet->CmdAccessory(746, 0, 1);
  XpressNet->CmdAccessory(746, 0, 0);
  XpressNet->CmdAccessory(746, 1, 1);
  XpressNet->CmdAccessory(746, 1, 0);
  XpressNet->CmdAccessory(751, 0, 1);
  XpressNet->CmdAccessory(751, 0, 0);

  XpressNet->CmdAccessory(751, 1, 1);
  XpressNet->CmdAccessory(751, 1, 0);
  XpressNet->CmdAccessory(753, 0, 1);
  XpressNet->CmdAccessory(753, 0, 0);
  XpressNet->CmdAccessory(754, 0, 1);
  XpressNet->CmdAccessory(754, 0, 0);
  XpressNet->CmdAccessory(745, 1, 1);
  XpressNet->CmdAccessory(746, 0, 0);
  XpressNet->CmdAccessory(751, 0, 1);
  XpressNet->CmdAccessory(753, 0, 0);

  XpressNet->CmdAccessory(745, 0, 1);
  XpressNet->CmdAccessory(745, 0, 0);
  XpressNet->CmdAccessory(745, 1, 1);
  XpressNet->CmdAccessory(745, 1, 0);
  XpressNet->CmdAccessory(746, 0, 1);
  XpressNet->CmdAccessory(746, 0, 0);
  XpressNet->CmdAccessory(746, 1, 1);
  XpressNet->CmdAccessory(746, 1, 0);
  XpressNet->CmdAccessory(751, 0, 1);
  XpressNet->CmdAccessory(751, 0, 0);
  XpressNet->CmdAccessory(751, 1, 1);

  XpressNet->CmdAccessory(751, 1, 0);
  XpressNet->CmdAccessory(753, 0, 1);
}

void FrmExtendedTest::CmdAllesAn_Click()
{
  XpressNet->CmdAllesAn();
}

void FrmExtendedTest::CmdAllesAus_Click()
{
  XpressNet->CmdAllesAus();
}

void FrmExtendedTest::CmdAlleLoksAnhalten_Click()
{
  XpressNet->CmdAlleLoksAnhalten();
}

void FrmExtendedTest::CmdEineLokAnhalten_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  XpressNet->CmdEineLokAnhalten(toInt(ui->txtLokAddress->text()));
}

void FrmExtendedTest::CmdLokEnter_Click()
{
  // A new lok address is entered. Get the lok info
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  ClearLedsForFunctions();
  XpressNet->CmdLokInfoAnfordern(toInt(ui->txtLokAddress->text()));
}

void FrmExtendedTest::CmdLokAusStack_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  XpressNet->CmdLokAusStack(toInt(ui->txtLokAddress->text()));
}

void FrmExtendedTest::CmdLokForward_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokSpeed->text())) return;
  XpressNet->CmdFahrbefehl(toInt(ui->txtLokAddress->text()), toInt(ui->TxtLokSpeed->text()), toInt(ui->ComboLokSteps->currentText()), true);
}

void FrmExtendedTest::CmdLokReverse_Click()
{
  if (!IsNumeric(ui->txtLokAddress->text())) return;
  if (!IsNumeric(ui->TxtLokSpeed->text())) return;
  XpressNet->CmdFahrbefehl(toInt(ui->txtLokAddress->text()), toInt(ui->TxtLokSpeed->text()), toInt(ui->ComboLokSteps->currentText()), false);
}

void FrmExtendedTest::CmdFunctionOn_Click()
{
  handleFunctionBefehl(true);
}

void FrmExtendedTest::CmdFunctionOff_Click()
{
  handleFunctionBefehl(false);
}

void FrmExtendedTest::CmdGetF0_F12_Click()
{
  XpressNet->CmdLokInfoAnfordern(toInt(ui->txtLokAddress->text()));
}

void FrmExtendedTest::CmdGetF13_28_Click()
{
  XpressNet->CmdFunktionsZustandF13F28Anfordern(toInt(ui->txtLokAddress->text()));
}

void FrmExtendedTest::CmdGetF29_68_Click()
{
  XpressNet->CmdFunktionsZustandF29F68Anfordern(toInt(ui->txtLokAddress->text()));
}


// ------------------------------------------------------------------------------------------------------------------
// Programming. Programming on Main as well as all three variants for Service Mode programming
// For simplicity, we do not check if the input values are valid
void FrmExtendedTest::CmdPomGet_Click()
{
  if (ui->CheckSM->isChecked()) {
    if (ui->SmDirect->isChecked()) XpressNet->CmdLeseanfrageProgrammieren(toInt(ui->TxtPomCV->text()));
    if (ui->SmRegister->isChecked()) XpressNet->CmdLeseanfrageProgrammierenRegistermode(toInt(ui->TxtPomCV->text()));
    if (ui->SmPage->isChecked()) XpressNet->CmdLeseanfrageProgrammierenPagemode(toInt(ui->TxtPomCV->text()));
  }
  else {
    XpressNet->CmdPomRead(toInt(ui->TxtPomAddress->text()), toInt(ui->TxtPomCV->text()));
  }
}

void FrmExtendedTest::CmdPomSet_Click()
{
  if (ui->CheckSM->isChecked()) {
    if (ui->SmDirect->isChecked()) XpressNet->CmdSchreibbefehlProgrammieren(toInt(ui->TxtPomCV->text()), toInt(ui->TxtPomValue->text()));
    if (ui->SmRegister->isChecked()) XpressNet->CmdSchreibbefehlProgrammierenRegisterMode(toInt(ui->TxtPomCV->text()), toInt(ui->TxtPomValue->text()));
    if (ui->SmPage->isChecked()) XpressNet->CmdSchreibbefehlProgrammierenPageMode(toInt(ui->TxtPomCV->text()), toInt(ui->TxtPomValue->text()));
  }
  else {
    XpressNet->CmdPomWrite(toInt(ui->TxtPomAddress->text()), toInt(ui->TxtPomCV->text()), toInt(ui->TxtPomValue->text()));
  }
}

void FrmExtendedTest::cmdPomHolen_Click()
{
  if (ui->CheckSM->isChecked()) {
    XpressNet->CmdProgrammierergebnisAnfordern();
  }
  else {
    XpressNet->CmdPoMErgebnisAnfordern();
  }
}

// ------------------------------------------------------------------------------------------------------------------
// Modellzeit
void FrmExtendedTest::CmdModellzeitAnfordern_Click()
{
  XpressNet->CmdModellzeitAnfordern();
}

void FrmExtendedTest::CmdModellzeitAnhalten_Click()
{
  XpressNet->CmdModellzeitAnhalten();
}

void FrmExtendedTest::CmdModellzeitStarten_Click()
{
  XpressNet->CmdModellzeitStarten();
}

void FrmExtendedTest::CmdModellzeitStellen_Click()
{
  quint8 day = 0;
  QString text = ui->ComboDay->currentText();
  if (text == "Monday") day = 0;
  if (text == "Tuesday") day = 1;
  if (text == "Wednesday") day = 2;
  if (text == "Thursday") day = 3;
  if (text == "Friday") day = 4;
  if (text == "Saterday") day = 5;
  if (text == "Sunday") day = 6;
  if (ui->CheckRun->isChecked() == false) ui->TxtFactor->setText("0");
  XpressNet->CmdModellzeitStellen(day, toInt(ui->TxtHour->text()), toInt(ui->TxtMinute->text()), toInt(ui->TxtFactor->text()));
}


// ==================================================================================================================
// EVENTS
void FrmExtendedTest::XpressNet_InterfaceConnected()
{
  ui->ledInterfaceStatus->setFillColor(Qt::green);
  ui->LabelStatus->setText("Connection established");
}

void FrmExtendedTest::XpressNet_TcpClosedByRemoteHost()
{
  ui->ledInterfaceStatus->setFillColor(Qt::red);
  ui->LabelStatus->setText("TCP connection closed by remote host");
}

void FrmExtendedTest::XpressNet_TcpError(QString ErrorDescription)
{
  ui->ledInterfaceStatus->setFillColor(Qt::red);
  ui->LabelStatus->setText(ErrorDescription);
}

void FrmExtendedTest::XpressNet_Feedback(int address, quint8 nibble, quint8 data)
{
  ui->TxBoxtRsAddress->setPlainText(QString::number(address) + "\n" + ui->TxBoxtRsAddress->toPlainText());
  ui->TxtBoxRsData->setPlainText(printBits(data) + "\n" + ui->TxtBoxRsData->toPlainText());
  ui->TxtBoxRsNibble->setPlainText(QString::number(nibble) + "\n" + ui->TxtBoxRsNibble->toPlainText());
}

void FrmExtendedTest::XpressNet_SwitchFeedback(int address, quint8 ttBits, quint8 iBit, quint8 data)
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

void FrmExtendedTest::XpressNet_AllesAn()
{
  ui->ledNotAus->setFillColor(Qt::green);
  ui->ledNotHalt->setFillColor(Qt::green);
  ui->ledProgrammiermode->setFillColor(Qt::green);
}

void FrmExtendedTest::XpressNet_NotAus()
{
  ui->ledNotAus->setFillColor(Qt::red);
}

void FrmExtendedTest::XpressNet_NotHalt()
{
  ui->ledNotHalt->setFillColor(Qt::red);
}

void FrmExtendedTest::XpressNet_ProgrammierMode()
{
  ui->ledProgrammiermode->setFillColor(Qt::red);
}

void FrmExtendedTest::XpressNet_ProgrammierinfoCV(int CV, quint8 Value)
{
  ui->TxtPomAddress->setText("");
  ui->TxtPomCV->setText(QString::number(CV));
  ui->TxtPomValue->setText(QString::number(Value));
}

void FrmExtendedTest::XpressNet_ProgrammierinfoPageRegister(quint8 EEPromAddress, quint8 Value)
{
  ui->TxtPomAddress->setText("");
  ui->TxtPomCV->setText(QString::number(EEPromAddress));
  ui->TxtPomValue->setText(QString::number(Value));
}

void FrmExtendedTest::XpressNet_ZentraleVersion(QString version)
{
  ui->txtMasterVersion->setText(version);
}

void FrmExtendedTest::XpressNet_ErweiterteZentraleVersion(QString BuildZentrale, QString VersionRuckMelder, QString BuildRuckMelder, QString VersionBootloader)
{
  ui->txtBuildNr->setText(BuildZentrale);
  ui->TxtVersionRuckmelder->setText(VersionRuckMelder);
  ui->txtBuildRuckmelder->setText(BuildRuckMelder);
  ui->txtVersionBootloader->setText(VersionBootloader);
}

void FrmExtendedTest::XpressNet_Fehlermeldung(quint8 Value)
{
  ui->LabelStatus->setText("Fehlermeldung: " + QString::number(Value) + " (V3.6: 3.1.14)");
}

void FrmExtendedTest::XpressNet_XpressNetAddress(quint8 address)
{
  ui->txtXAddress->setText(QString::number(address));
}

void FrmExtendedTest::XpressNet_XpressNetVersion(QString version)
{
  ui->txtXVersion->setText(version);
}

void FrmExtendedTest::XpressNet_InterfaceVersion(QString version, QString code)
{
  ui->txtInterfaceVersion->setText(version);
  ui->txtInterfaceCode->setText(code);
}

void FrmExtendedTest::XpressNet_InterfaceFreieVerbindungen(quint8 number)
{
  ui->txtTcpFree->setText(QString::number(number));
}

void FrmExtendedTest::XpressNet_Modellzeit(quint8 day, quint8 hour, quint8 minute, quint8 Faktor)
{
  Q_UNUSED(day);
  Q_UNUSED(Faktor);
  ui->TxtHours->setText(QString::number(hour));
  if (minute < 10) {
    ui->TxtMinutes->setText("0" + QString::number(minute));
  }
  else {
    ui->TxtMinutes->setText(QString::number(minute));
  }
}

void FrmExtendedTest::XpressNet_NormaleLokinfo(bool Besetzt, int Speedsteps, int Speed, bool Forward, quint8 F0_F4, quint8 F5_F8, quint8 F9_F12)
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

void FrmExtendedTest::XpressNet_FunctionsZustandF13F20(quint8 F13_F20, quint8 F21_F28)
{
  FB[4] = F13_F20;                                        // Store Gruppe 4
  FB[5] = F21_F28;                                        // Store Gruppe 5
  ShowFbGroup(F13_F20, 4);                                // Led FB bits for Gruppe 4
  ShowFbGroup(F21_F28, 5);                                // Led FB bits for Gruppe 5
}

void FrmExtendedTest::XpressNet_FunctionsZustandF29F68(quint8 F29_F36, quint8 F37_F44, quint8 F45_F52, quint8 F53_F60, quint8 F61_F68)
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

void FrmExtendedTest::TimerMessageCounter_Timer()
{
  ui->txtNmbPckts->setText(QString::number(XpressNet->NumberOfExpressNetMessagesReceived + XpressNet->NumberOfExpressNetMessagesTransmitted));
}

// Version 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum
void FrmExtendedTest::XpressNet_SendError(quint8 Reason, QString Description)
{
  ui->LabelStatus->setText(sendErrorName(Reason) + ": " + Description);
}

// Version 2.0: the command station does not address the interface: sending is paused
void FrmExtendedTest::XpressNet_InterfaceNichtAdressiert()
{
  ui->LabelStatus->setText("Interface nicht adressiert");
}

// Version 2.0: the command station addresses the interface again
void FrmExtendedTest::XpressNet_InterfaceWiederAdressiert()
{
  ui->LabelStatus->setText("Interface wieder adressiert");
}


// ==================================================================================================================
// Support routines below
// ------------------------------------------------------------------------------------------------------------------
// The VB6 control array LEDF0F68: the Led widgets LEDF0F68_0 .. LEDF0F68_68 of the form
void FrmExtendedTest::InitControlArrays()
{
  for (int i = 0; i <= 68; i++) LEDF0F68[i] = findChild<Led *>("LEDF0F68_" + QString::number(i));
}

// ------------------------------------------------------------------------------------------------------------------
// Lok Functions
void FrmExtendedTest::handleFunctionBefehl(bool turnOn)
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

void FrmExtendedTest::ClearLedsForFunctions()
{
  for (int i = 0; i <= 68; i++) {
    LEDF0F68[i]->setFillColor(Qt::white);
  }
}

QColor FrmExtendedTest::FbColor(bool Value)
{
  if (Value) return Qt::red;
  else return Qt::green;
}

// ------------------------------------------------------------------------------------------------------------------
// Functions to show the Feedbacks on this form
void FrmExtendedTest::ShowFbGroup(quint8 data, quint8 group)
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
bool FrmExtendedTest::bitValue(quint8 data, quint8 index)
{
  return ((data >> index) & 1) == 1;
}

quint8 FrmExtendedTest::bitSet(quint8 data, quint8 index)
{
  return quint8(data | (1 << index));
}

quint8 FrmExtendedTest::bitClear(quint8 data, quint8 index)
{
  return quint8(data & (255 - (1 << index)));            // mask the data
}

quint8 FrmExtendedTest::bitUpdate(quint8 data, quint8 index, bool bitIsSet)
{
  if (bitIsSet) return bitSet(data, index);
  else return bitClear(data, index);
}

// ------------------------------------------------------------------------------------------------------------------
// Transforms numbers form 0..15 (4 bits) in a bitstring
QString FrmExtendedTest::printBits(quint8 data)
{
  QString outString;
  for (int i = 0; i <= 3; i++) {
    outString = QString::number(data % 2) + outString;
    data = data / 2;
  }
  return outString;
}
