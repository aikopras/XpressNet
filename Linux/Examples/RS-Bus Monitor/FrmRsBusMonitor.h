//
//  FrmRsBusMonitor.h
//  XpressNetQt examples: RS-Bus Monitor
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See FrmRsBusMonitor.cpp
//

#ifndef FRMRSBUSMONITOR_H
#define FRMRSBUSMONITOR_H

#include <QWidget>
#include "XpressNet.h"

class PanelManager;
namespace Ui { class FrmRsBusMonitor; }

class FrmRsBusMonitor : public QWidget
{
  Q_OBJECT

public:
  explicit FrmRsBusMonitor(QWidget *parent = nullptr);
  ~FrmRsBusMonitor() override;

protected:
  void resizeEvent(QResizeEvent *event) override;         // .NET: FrmRsBusMonitor_Resize

private slots:
  void cmdTcpConnect_Click();
  void CmdAllesAn_Click();
  void CmdAllesAus_Click();
  void RS2Switch_Click();
  void Switch2RS_Click();
  void SchaltInfo_Click();
  // Events of the XpressNet object
  void XpressNet_InterfaceConnected();
  void XpressNet_TcpClosedByRemoteHost();
  void XpressNet_TcpError(QString ErrorDescription);
  void XpressNet_AllesAn();
  void XpressNet_NotAus();
  void XpressNet_NotHalt();
  void XpressNet_SendError(quint8 Reason, QString Description);
  void XpressNet_InterfaceNichtAdressiert();
  void XpressNet_InterfaceWiederAdressiert();
  void XpressNet_Feedback(int address, quint8 Nibble, quint8 data);
  void XpressNet_SwitchFeedback(int switchAddress, quint8 ttBits, quint8 iBit, quint8 data);

private:
  Ui::FrmRsBusMonitor *ui;
  XpressNetClass *XpressNet;

  // Referentie naar de PanelManager
  PanelManager *m_PanelManager;

  static const int LEFT_OFFSET = 240;

  bool GetSchaltInfo;
  int AnswerCount;                                        // The answers for RS-Bus address 64 and 128

  void RequestFeedbacks(int firstAddress, int lastAddress);
  void CheckSchaltInfo(int RsAddress);
};

#endif // FRMRSBUSMONITOR_H
