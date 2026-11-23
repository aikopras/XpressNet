//
//  PreferencesDialog.h
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The preferences window (macOS: PreferencesController.xib). The layout is created in code. The values are stored
//  with OK (macOS: directly, via bindings); they are used the next time SCAN is pushed.
//

#ifndef PREFERENCESDIALOG_H
#define PREFERENCESDIALOG_H

#include <QDialog>
#include <QList>

class QLineEdit;
class QCheckBox;

class PreferencesDialog : public QDialog
{
  Q_OBJECT

public:
  explicit PreferencesDialog(QWidget *parent = nullptr);

signals:
  void cvOverviewRequested();                             // The button "CV overview..."

private slots:
  void save();

private:
  QList<QLineEdit *> ipFields;
  QList<QCheckBox *> cvBoxes;                             // One per common CV (commonCvs in Settings.h)
  QLineEdit *cvOther;
  QCheckBox *liftEnabled;
  QLineEdit *liftSwitchAddress;
  QLineEdit *liftIrRsAddress;
};

#endif // PREFERENCESDIALOG_H
