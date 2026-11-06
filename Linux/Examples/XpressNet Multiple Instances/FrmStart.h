//
//  FrmStart.h
//  XpressNetQt examples: XpressNet Multiple Instances
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See FrmStart.cpp
//

#ifndef FRMSTART_H
#define FRMSTART_H

#include <QList>
#include <QPointer>
#include <QWidget>

class FrmExtendedTest;
class FrmZ21Test;
namespace Ui { class FrmStart; }

class FrmStart : public QWidget
{
  Q_OBJECT

public:
  explicit FrmStart(QWidget *parent = nullptr);
  ~FrmStart() override;

protected:
  void closeEvent(QCloseEvent *event) override;

private slots:
  void cmdNewXpressNet_Click();
  void cmdNewZ21_Click();

private:
  Ui::FrmStart *ui;
  QList<QPointer<FrmExtendedTest>> XpnConnections;       // QPointer: becomes nullptr when the form is closed
  QList<QPointer<FrmZ21Test>> Z21Connections;
};

#endif // FRMSTART_H
