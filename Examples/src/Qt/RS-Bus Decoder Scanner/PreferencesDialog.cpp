//
//  PreferencesDialog.cpp
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See PreferencesDialog.h
//

#include "PreferencesDialog.h"
#include "Settings.h"
#include "DecoderScanner.h"
#include <QCheckBox>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QGridLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QVBoxLayout>

// A small grey note below a field
static QLabel *note(const QString &text)
{
  QLabel *label = new QLabel(text);
  label->setWordWrap(true);
  QFont font = label->font();
  font.setPointSizeF(font.pointSizeF() * 0.9);
  label->setFont(font);
  label->setStyleSheet("color: gray;");
  return label;
}

PreferencesDialog::PreferencesDialog(QWidget *parent) : QDialog(parent)
{
  setWindowTitle("RS-Bus Decoder Scanner - Preferences");
  QSettings &s = Settings::settings();
  QVBoxLayout *layout = new QVBoxLayout(this);
  // The IP addresses of the command stations
  QFormLayout *ipLayout = new QFormLayout;
  for (int number = 1; number <= MAX_STATIONS; number++) {
    QLineEdit *field = new QLineEdit(s.value(Settings::ipKey(number)).toString());
    field->setMinimumWidth(200);
    ipLayout->addRow(QString("IP address command station %1:").arg(number), field);
    ipFields.append(field);
  }
  layout->addLayout(ipLayout);
  layout->addWidget(note("Leave the field empty if the command station is not used. Factory default: 192.168.0.200"));
  // The CVs in the table: the common CVs in two columns, then the other CVs
  layout->addSpacing(12);
  layout->addWidget(new QLabel("CVs in the table:"));
  QGridLayout *cvLayout = new QGridLayout;
  int rows = (numberOfCommonCvs + 1) / 2;
  for (int i = 0; i < numberOfCommonCvs; i++) {
    QString title = QString("CV%1 %2%3").arg(commonCvs[i].cv).arg(commonCvs[i].name).arg(commonCvs[i].newOnly ? " *" : "");
    QCheckBox *box = new QCheckBox(title);
    box->setChecked(s.value(Settings::checkedKey(commonCvs[i].cv), false).toBool());
    cvLayout->addWidget(box, i % rows, i / rows);
    cvBoxes.append(box);
  }
  cvLayout->setContentsMargins(20, 0, 0, 0);
  layout->addLayout(cvLayout);
  layout->addWidget(note("* Version 20 or higher only"));
  QFormLayout *otherLayout = new QFormLayout;
  cvOther = new QLineEdit(s.value("cvOther").toString());
  otherLayout->addRow("Other CVs:", cvOther);
  layout->addLayout(otherLayout);
  layout->addWidget(note("For example: 33, 34, 64. In the table the columns can be moved and resized.\n"
                         "The new values are used the next time SCAN is pushed."));
  // The lift decoder
  layout->addSpacing(12);
  QGridLayout *liftLayout = new QGridLayout;
  liftEnabled = new QCheckBox("Include lift decoder, first switch address:");
  liftEnabled->setChecked(s.value("liftEnabled", false).toBool());
  liftSwitchAddress = new QLineEdit(s.value("liftSwitchAddress").toString());
  liftSwitchAddress->setMaximumWidth(80);
  liftIrRsAddress = new QLineEdit(s.value("liftIrRsAddress").toString());
  liftIrRsAddress->setMaximumWidth(80);
  liftLayout->addWidget(liftEnabled, 0, 0);
  liftLayout->addWidget(liftSwitchAddress, 0, 1);
  liftLayout->addWidget(note("The PoM address of the lift decoder is (switch address - 1) / 4 + 7000."), 1, 0, 1, 2);
  QLabel *irLabel = new QLabel("RS-Bus address of the IR sensors of the lift:");
  irLabel->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
  liftLayout->addWidget(irLabel, 2, 0);
  liftLayout->addWidget(liftIrRsAddress, 2, 1);
  liftLayout->addWidget(note("The IR sensors use this address and the next one. Empty: no IR sensors."), 3, 0, 1, 2);
  layout->addLayout(liftLayout);
  // CV overview, OK and Cancel
  layout->addSpacing(12);
  QDialogButtonBox *buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel);
  QPushButton *overview = buttons->addButton("CV overview...", QDialogButtonBox::HelpRole);
  connect(overview, &QPushButton::clicked, this, &PreferencesDialog::cvOverviewRequested);
  connect(buttons, &QDialogButtonBox::accepted, this, &PreferencesDialog::save);
  connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
  layout->addWidget(buttons);
}

void PreferencesDialog::save()
{
  QSettings &s = Settings::settings();
  for (int i = 0; i < ipFields.count(); i++) s.setValue(Settings::ipKey(i + 1), ipFields[i]->text().trimmed());
  for (int i = 0; i < cvBoxes.count(); i++) s.setValue(Settings::checkedKey(commonCvs[i].cv), cvBoxes[i]->isChecked());
  s.setValue("cvOther", cvOther->text());
  s.setValue("liftEnabled", liftEnabled->isChecked());
  s.setValue("liftSwitchAddress", liftSwitchAddress->text().trimmed());
  s.setValue("liftIrRsAddress", liftIrRsAddress->text().trimmed());
  accept();
}
