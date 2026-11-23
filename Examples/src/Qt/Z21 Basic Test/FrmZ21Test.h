//
//  FrmZ21Test.h
//  XpressNetQt examples: Z21 Basic Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See FrmZ21Test.cpp
//

#ifndef FRMZ21TEST_H
#define FRMZ21TEST_H

#include <QWidget>
#include <QTimer>
#include "XpressNet.h"

class Led;
namespace Ui { class FrmZ21Test; }

class FrmZ21Test : public QWidget
{
  Q_OBJECT

public:
  explicit FrmZ21Test(QWidget *parent = nullptr);
  ~FrmZ21Test() override;

private slots:
  void TimerMessageCounter_Timer();
  void cmdUdpSet_Click();
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
  void CmdGetLokInfo_Click();
  void cmdRecht_Click();
  void cmdAf_Click();
  void cmdSwitchFeedback_Click();
  void CmdExtAccSet_Click();
  // Events of the XpressNet object
  void XpressNet_SwitchFeedback(int address, quint8 ttBits, quint8 iBit, quint8 data);
  void XpressNet_AllesAn();
  void XpressNet_NotAus();
  void XpressNet_NotHalt();
  void XpressNet_ProgrammierMode();
  void XpressNet_ZentraleVersion(QString version);
  void XpressNet_Z21FirmwareVersion(QString version);
  void XpressNet_Z21Lokinfo(int LokAddress, bool Besetzt, int Speedsteps, int Speed, bool Forward, quint8 F0_F4, quint8 F5_F8, quint8 F9_F12, quint8 F13_F20, quint8 F21_F28);
  void XpressNet_SendError(quint8 Reason, QString Description);

private:
  Ui::FrmZ21Test *ui;
  XpressNetClass *XpressNet;
  QTimer TimerMessageCounter;                             // .NET: a Timer component of the form (1000 ms)

  // Function Groups:
  // For each Lok Address, the values must be fethed again
  quint8 FB[11];                                          // FB: Function Befehl (on / off)
  quint8 FS[11];                                          // Function Status (tastend oder nicht tastend)

  // The control array of the VB6 form (see InitControlArrays)
  Led *LEDF0F68[29];

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

#endif // FRMZ21TEST_H
