//
//  FrmBasicTest.h
//  XpressNetQt examples: XpressNet Basic Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See FrmBasicTest.cpp
//

#ifndef FRMBASICTEST_H
#define FRMBASICTEST_H

#include <QWidget>
#include "XpressNet.h"

class Led;
namespace Ui { class FrmBasicTest; }

class FrmBasicTest : public QWidget
{
  Q_OBJECT

public:
  explicit FrmBasicTest(QWidget *parent = nullptr);
  ~FrmBasicTest() override;

protected:
  bool eventFilter(QObject *watched, QEvent *event) override;

private slots:
  void cmdTcpConnect_Click();
  void CmdUsbConnect_Click();
  void LabelMasterVersion_Click();
  void CmdInterfaceInfo_Click();
  void CmdLanInfo_Click();
  void cmdRecht_Click();
  void cmdAf_Click();
  void cmdSwitchFeedback_Click();
  void CmdRsFeedback_Click();
  void CmdAllesAn_Click();
  void CmdAllesAus_Click();
  void CmdAlleLoksAnhalten_Click();
  void CmdEineLokAnhalten_Click();
  void CmdLokEnter_Click();
  void CmdLokAusStack_Click();
  void CmdLokForward_Click();
  void CmdLokReverse_Click();
  void CmdFunctionOn_Click();
  void CmdFunctionOff_Click();
  void CmdGetF0_F12_Click();
  void CmdGetF13_28_Click();
  void CmdGetF29_68_Click();
  // Events of the XpressNet object
  void XpressNet_InterfaceConnected();
  void XpressNet_TcpClosedByRemoteHost();
  void XpressNet_TcpError(QString ErrorDescription);
  void XpressNet_Feedback(int address, quint8 nibble, quint8 data);
  void XpressNet_SwitchFeedback(int address, quint8 ttBits, quint8 iBit, quint8 data);
  void XpressNet_AllesAn();
  void XpressNet_NotAus();
  void XpressNet_NotHalt();
  void XpressNet_ProgrammierMode();
  void XpressNet_ZentraleVersion(QString version);
  void XpressNet_Fehlermeldung(quint8 Value);
  void XpressNet_XpressNetAddress(quint8 address);
  void XpressNet_XpressNetVersion(QString version);
  void XpressNet_InterfaceVersion(QString version, QString code);
  void XpressNet_InterfaceFreieVerbindungen(quint8 number);
  void XpressNet_NormaleLokinfo(bool Besetzt, int Speedsteps, int Speed, bool Forward, quint8 F0_F4, quint8 F5_F8, quint8 F9_F12);
  void XpressNet_FunctionsZustandF13F20(quint8 F13_F20, quint8 F21_F28);
  void XpressNet_FunctionsZustandF29F68(quint8 F29_F36, quint8 F37_F44, quint8 F45_F52, quint8 F53_F60, quint8 F61_F68);
  void XpressNet_SendError(quint8 Reason, QString Description);
  void XpressNet_InterfaceNichtAdressiert();
  void XpressNet_InterfaceWiederAdressiert();

private:
  Ui::FrmBasicTest *ui;
  XpressNetClass *XpressNet;

  // Function Groups:
  // For each Lok Address, the values must be fethed again
  quint8 FB[11];                                          // FB: Function Befehl (on / off)
  quint8 FS[11];                                          // Function Status (tastend oder nicht tastend)

  // The control array of the VB6 form (see InitControlArrays)
  Led *LEDF0F68[69];

  void InitControlArrays();
  void handleFunctionBefehl(bool turnOn);
  void ClearLedsForFunctions();
  static QColor FbColor(bool Value);
  void ShowFbGroup(quint8 data, quint8 group);
  static bool bitValue(quint8 data, quint8 index);
  static quint8 bitSet(quint8 data, quint8 index);
  static quint8 bitClear(quint8 data, quint8 index);
  static quint8 bitUpdate(quint8 data, quint8 index, bool bitIsSet);
  static QString printBits(quint8 data);
};

#endif // FRMBASICTEST_H
