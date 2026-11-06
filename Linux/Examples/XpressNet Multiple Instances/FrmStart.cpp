//
//  FrmStart.cpp
//  XpressNetQt examples: XpressNet Multiple Instances
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of the example XpressNet Multiple Instances (VB6: B_UserInterfaceFormStart.frm, .NET: FrmStart.vb), for
//  XpressNetQt 2.0. Every click opens a new form with its own XpressNetClass object: the Extended test
//  (FrmExtendedTest) or the Z21 Basic test (FrmZ21Test). These forms are the same as in the examples XpressNet Test
//  and Z21 Basic Test. In the logfile, every XpressNet object has its own letter (A, B, ...).
//  Differences with .NET:
//  - A closed form is deleted (Qt::WA_DeleteOnClose); its destructor deletes its XpressNet object (.NET: Dispose).
//  - Closing this form closes the other forms as well (.NET: the end of the application).
//

#include "FrmStart.h"
#include "ui_FrmStart.h"
#include "FrmExtendedTest.h"
#include "FrmZ21Test.h"
#include <QCloseEvent>


FrmStart::FrmStart(QWidget *parent) : QWidget(parent), ui(new Ui::FrmStart)
{
  ui->setupUi(this);
  connect(ui->cmdNewXpressNet, &QPushButton::clicked, this, &FrmStart::cmdNewXpressNet_Click);
  connect(ui->cmdNewZ21, &QPushButton::clicked, this, &FrmStart::cmdNewZ21_Click);
}

FrmStart::~FrmStart()
{
  delete ui;
}

// Closing this form closes the other forms as well. These are deleted immediately, so their XpressNet objects close
// the connections before the application ends
void FrmStart::closeEvent(QCloseEvent *event)
{
  for (const QPointer<FrmExtendedTest> &XpnConnection : XpnConnections) delete XpnConnection.data();
  for (const QPointer<FrmZ21Test> &Z21Connection : Z21Connections) delete Z21Connection.data();
  XpnConnections.clear();
  Z21Connections.clear();
  event->accept();
}

void FrmStart::cmdNewXpressNet_Click()
{
  FrmExtendedTest *XpnConnection = new FrmExtendedTest();
  XpnConnection->setAttribute(Qt::WA_DeleteOnClose);
  XpnConnections.append(XpnConnection);
  XpnConnection->show();
}

void FrmStart::cmdNewZ21_Click()
{
  FrmZ21Test *Z21Connection = new FrmZ21Test();
  Z21Connection->setAttribute(Qt::WA_DeleteOnClose);
  Z21Connections.append(Z21Connection);
  Z21Connection->show();
}
