//
//  FrmRsBusMonitor.cpp
//  XpressNetQt examples: RS-Bus Monitor
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of the program RS-Bus Monitor (VB6: RsBusMonitor.frm, .NET: FrmRsBusMonitor.vb), for XpressNetQt 2.0.
//  The layout (FrmRsBusMonitor.ui) is converted from the .NET form; the code follows the .NET code.
//  Differences with .NET:
//  - The events of the XpressNet object are Qt signals, connected in the constructor (.NET: Handles).
//  - The LEDs are Led widgets (Led.cpp). picPanel is a QWidget, in which the PanelManager draws (PanelManager.cpp).
//  - The IP address is kept with QSettings, in ~/.config/RS-Bus Monitor/Settings.conf (.NET: SaveSetting / GetSetting).
//  - The logfile directory is ~/TMC-Logfiles (.NET: C:\Users\Public\TMC-Logfiles).
//  - FormClosed / XpressNet.Dispose() is the destructor (delete XpressNet).
//  Get RS-Bus data stops after SendError, a new connection, TcpClosedByRemoteHost and TcpError.
//

#include "FrmRsBusMonitor.h"
#include "ui_FrmRsBusMonitor.h"
#include "Led.h"
#include "PanelManager.h"
#include "XpressNetLogFile.h"
#include <QDir>
#include <QSettings>

// VB: IsNumeric. Is the text a number?
static bool IsNumeric(const QString &text)
{
  bool ok = false;
  text.trimmed().toDouble(&ok);
  return ok;
}

// VB: CInt. The conversion of a text to a number
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


FrmRsBusMonitor::FrmRsBusMonitor(QWidget *parent) : QWidget(parent), ui(new Ui::FrmRsBusMonitor)
{
  m_PanelManager = nullptr;
  ui->setupUi(this);
  XpressNet = new XpressNetClass(this);
  // Set the directory for logfiles (on Linux the default directory of the library)
  XpressNet->Logfile->setLogDirectory(QDir(QDir::homePath()).filePath("TMC-Logfiles"));
  // The IP address of the previous connection (the first time: the text of txtIP)
  QSettings settings("RS-Bus Monitor", "Settings");
  ui->txtIP->setText(settings.value("IpAddress", ui->txtIP->text()).toString());
  ui->picPanel->setGeometry(LEFT_OFFSET, 0, width() - LEFT_OFFSET, height());
  // Initialiseer de PanelManager
  m_PanelManager = new PanelManager(this);
  m_PanelManager->Init(ui->picPanel);
  GetSchaltInfo = false;
  AnswerCount = 0;
  // The buttons
  connect(ui->cmdTcpConnect, &QPushButton::clicked, this, &FrmRsBusMonitor::cmdTcpConnect_Click);
  connect(ui->CmdAllesAn, &QPushButton::clicked, this, &FrmRsBusMonitor::CmdAllesAn_Click);
  connect(ui->CmdAllesAus, &QPushButton::clicked, this, &FrmRsBusMonitor::CmdAllesAus_Click);
  connect(ui->RS2Switch, &QPushButton::clicked, this, &FrmRsBusMonitor::RS2Switch_Click);
  connect(ui->Switch2RS, &QPushButton::clicked, this, &FrmRsBusMonitor::Switch2RS_Click);
  connect(ui->SchaltInfo, &QPushButton::clicked, this, &FrmRsBusMonitor::SchaltInfo_Click);
  // The events of the XpressNet object
  connect(XpressNet, &XpressNetClass::InterfaceConnected, this, &FrmRsBusMonitor::XpressNet_InterfaceConnected);
  connect(XpressNet, &XpressNetClass::TcpClosedByRemoteHost, this, &FrmRsBusMonitor::XpressNet_TcpClosedByRemoteHost);
  connect(XpressNet, &XpressNetClass::TcpError, this, &FrmRsBusMonitor::XpressNet_TcpError);
  connect(XpressNet, &XpressNetClass::AllesAn, this, &FrmRsBusMonitor::XpressNet_AllesAn);
  connect(XpressNet, &XpressNetClass::NotAus, this, &FrmRsBusMonitor::XpressNet_NotAus);
  connect(XpressNet, &XpressNetClass::NotHalt, this, &FrmRsBusMonitor::XpressNet_NotHalt);
  connect(XpressNet, &XpressNetClass::SendError, this, &FrmRsBusMonitor::XpressNet_SendError);
  connect(XpressNet, &XpressNetClass::InterfaceNichtAdressiert, this, &FrmRsBusMonitor::XpressNet_InterfaceNichtAdressiert);
  connect(XpressNet, &XpressNetClass::InterfaceWiederAdressiert, this, &FrmRsBusMonitor::XpressNet_InterfaceWiederAdressiert);
  connect(XpressNet, &XpressNetClass::Feedback, this, &FrmRsBusMonitor::XpressNet_Feedback);
  connect(XpressNet, &XpressNetClass::SwitchFeedback, this, &FrmRsBusMonitor::XpressNet_SwitchFeedback);
}

// .NET: FormClosed / XpressNet.Dispose(). Release the XpressNet object (which in turn closes the TCP connection)
FrmRsBusMonitor::~FrmRsBusMonitor()
{
  delete XpressNet;
  delete ui;
}

void FrmRsBusMonitor::resizeEvent(QResizeEvent *event)
{
  QWidget::resizeEvent(event);
  if (m_PanelManager == nullptr) return;                  // The form is not yet loaded
  if (!isMinimized()) {
    ui->picPanel->setGeometry(LEFT_OFFSET, 0, width() - LEFT_OFFSET, height());
    m_PanelManager->Resize();
  }
}


// ==================================================================================================================
// XPressNet Commando's
void FrmRsBusMonitor::cmdTcpConnect_Click()
{
  ui->LabelStatus->setText("TCP connecting .....");
  // The IP address is kept for the next start of the program
  QSettings settings("RS-Bus Monitor", "Settings");
  settings.setValue("IpAddress", ui->txtIP->text());
  // A new connection stops Get RS-Bus data: the answers on its requests will not come
  GetSchaltInfo = false;
  XpressNet->CmdConnectTcp(ui->txtIP->text(), "5550");
}

void FrmRsBusMonitor::CmdAllesAn_Click()
{
  XpressNet->CmdAllesAn();
}

void FrmRsBusMonitor::CmdAllesAus_Click()
{
  XpressNet->CmdAllesAus();
}

// ==================================================================================================================
// ALGEMENE EVENTS
void FrmRsBusMonitor::XpressNet_InterfaceConnected()
{
  ui->ledInterfaceStatus->setFillColor(Qt::green);
  ui->LabelStatus->setText("Connection established");
}

void FrmRsBusMonitor::XpressNet_TcpClosedByRemoteHost()
{
  ui->ledInterfaceStatus->setFillColor(Qt::red);
  GetSchaltInfo = false;                                  // Get RS-Bus data stops
  ui->LabelStatus->setText("TCP connection closed by remote host");
}

void FrmRsBusMonitor::XpressNet_TcpError(QString ErrorDescription)
{
  ui->ledInterfaceStatus->setFillColor(Qt::red);
  GetSchaltInfo = false;                                  // Get RS-Bus data stops
  ui->LabelStatus->setText(ErrorDescription);
}

void FrmRsBusMonitor::XpressNet_AllesAn()
{
  ui->ledNotAus->setFillColor(Qt::green);
  ui->ledNotHalt->setFillColor(Qt::green);
}

void FrmRsBusMonitor::XpressNet_NotAus()
{
  ui->ledNotAus->setFillColor(Qt::red);
}

void FrmRsBusMonitor::XpressNet_NotHalt()
{
  ui->ledNotHalt->setFillColor(Qt::red);
}

// Version 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum
// Get RS-Bus data is stopped (for example: no answer, the message queues of the library have been emptied)
void FrmRsBusMonitor::XpressNet_SendError(quint8 Reason, QString Description)
{
  GetSchaltInfo = false;
  ui->LabelStatus->setText(sendErrorName(Reason) + ": " + Description);
}

// Version 2.0: the command station does not address the interface: sending is paused
void FrmRsBusMonitor::XpressNet_InterfaceNichtAdressiert()
{
  ui->LabelStatus->setText("Interface nicht adressiert");
}

// Version 2.0: the command station addresses the interface again
void FrmRsBusMonitor::XpressNet_InterfaceWiederAdressiert()
{
  ui->LabelStatus->setText("Interface wieder adressiert");
}


// ==================================================================================================================
// FEEDBACK EVENTS

// We hebben in de library 2 events voor RS-Bus input: Bezetmeld en Wissel feedbacks
// Voor beide roepen we SetFeedbackNibble aan
//
// Bezetmeld: XpressNet_Feedback(int address, quint8 nibble, quint8 data)
// - address = 1..128
// - nibble = 0 / 1
// - data = 0..15
void FrmRsBusMonitor::XpressNet_Feedback(int address, quint8 Nibble, quint8 data)
{
  m_PanelManager->SetFeedbackNibble(address, Nibble, data, BlockStateType::BEZETMELDING);
  // Get RS-Bus data: is this the answer on the last request of a part?
  CheckSchaltInfo(address);
}
//
// Wissel: XpressNet_SwitchFeedback(switchAddress, ttBits, iBit, data)
// - switchAddress =  ((address * 4) + (Nibble * 2) + 1)
// - ttBits =
// - iBit =
// - data = 0..15
void FrmRsBusMonitor::XpressNet_SwitchFeedback(int switchAddress, quint8 ttBits, quint8 iBit, quint8 data)
{
  Q_UNUSED(iBit);
  int RsAddress;
  quint8 Nibble;
  RsAddress = (switchAddress - 1) / 4 + 1;
  Nibble = ((switchAddress - 1) / 2) % 2;
  if (ttBits == 0) {     // Feedback comes from Command Station
    if (!GetSchaltInfo) {
      m_PanelManager->SetFeedbackNibble(RsAddress, Nibble, data, BlockStateType::CENTRALE);
    }
  } else {               // Feedback comes from Decoder
    m_PanelManager->SetFeedbackNibble(RsAddress, Nibble, data, BlockStateType::WISSEL);
  }
  // Get RS-Bus data: is this the answer on the last request of a part?
  CheckSchaltInfo(RsAddress);
}


// ==================================================================================================================
// VERTALING TUSSE ADDRESSEN
void FrmRsBusMonitor::RS2Switch_Click()
{
  int rsAddr;
  int firstSw;
  int lastSw;
  if (IsNumeric(ui->TextRS2Switch->text())) {
    rsAddr = toInt(ui->TextRS2Switch->text());
    if (rsAddr >= 1) {
      firstSw = (rsAddr - 1) * 4 + 1;
      lastSw = firstSw + 3;
      ui->TextSwitch2RS->setText(QString::number(firstSw) + " .. " + QString::number(lastSw));
    }
  }
  // Ongeldige input? Doe niets
}

void FrmRsBusMonitor::Switch2RS_Click()
{
  int swAddr;
  int rsAddr;
  if (IsNumeric(ui->TextSwitch2RS->text())) {
    swAddr = toInt(ui->TextSwitch2RS->text());
    if (swAddr >= 1) {
      rsAddr = (swAddr - 1) / 4 + 1;
      ui->TextRS2Switch->setText(QString::number(rsAddr));
    }
  }
  // Ongeldige input? Doe niets
}


// ==================================================================================================================
// Alles ophalen
// Get RS-Bus data in two parts of 128 requests (64 RS-Bus addresses, two nibbles each). The FIFO of version 2.0 of
// the library holds 128 messages: more requests at once would give the event SendError (FifoVoll).
// The second part is requested as soon as both requests for RS-Bus address 64 have been answered; Get RS-Bus data
// ends when both requests for RS-Bus address 128 have been answered (see CheckSchaltInfo). This does not depend on
// the speed of the interface. The event SendError stops Get RS-Bus data.
void FrmRsBusMonitor::SchaltInfo_Click()
{
  if (GetSchaltInfo) return;                              // Get RS-Bus data is already active
  GetSchaltInfo = true;
  AnswerCount = 0;
  ui->LabelStatus->setText("Requesting RS-Bus data .....");
  RequestFeedbacks(1, 64);
}

// Requests both nibbles of the RS-Bus addresses firstAddress .. lastAddress
void FrmRsBusMonitor::RequestFeedbacks(int firstAddress, int lastAddress)
{
  for (int address = firstAddress; address <= lastAddress; address++) {
    XpressNet->CmdFeedbackRequest(address, 0);            // Nibble = 0
    XpressNet->CmdFeedbackRequest(address, 1);            // Nibble = 1
  }
}

// Called for every feedback message, after the matrix has been updated. The answers are counted per RS-Bus address,
// and not per nibble: for an RS-Bus address without feedback information, the command station (LZ100) answers both
// requests with nibble 0 (42 Address 00).
// Note: an unrequested feedback message (broadcast) for RS-Bus address 64 or 128 is counted as well
void FrmRsBusMonitor::CheckSchaltInfo(int RsAddress)
{
  if (!GetSchaltInfo) return;
  if ((RsAddress != 64) && (RsAddress != 128)) return;
  AnswerCount = AnswerCount + 1;
  if (AnswerCount == 2) RequestFeedbacks(65, 128);        // The first part has been answered
  if (AnswerCount == 4) {                                 // The second part has been answered
    GetSchaltInfo = false;
    ui->LabelStatus->setText("RS-Bus data received");
  }
}
