//
//  Settings.cpp
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See Settings.h
//

#include "Settings.h"
#include "DecoderScanner.h"
#include <QRegularExpression>

QSettings &Settings::settings()
{
  static QSettings settings("RS-Bus Decoder Scanner", "Settings");
  return settings;
}

void Settings::checkDefaults()
{
  QSettings &s = settings();
  // The IP addresses: the default address of the Lenz interface for the first command station
  for (int number = 1; number <= MAX_STATIONS; number++) {
    if (!s.contains(ipKey(number))) s.setValue(ipKey(number), (number == 1) ? "192.168.0.200" : "");
  }
  if (!s.contains("cvOther")) {
    for (int cv : defaultCvs) s.setValue(checkedKey(cv), true);
    s.setValue("cvOther", "");
  }
  // The lift decoder: the defaults of mySettings.h of Lift_Main (CV1 = 56, CV9 = 3, thus switch address 989) and of
  // Lift_IR (RS-Bus address 124)
  if (!s.contains("liftEnabled")) s.setValue("liftEnabled", false);
  if (!s.contains("liftSwitchAddress")) s.setValue("liftSwitchAddress", "989");
  if (!s.contains("liftIrRsAddress")) s.setValue("liftIrRsAddress", "124");
}

QStringList Settings::ipAddresses()
{
  QStringList list;
  for (int number = 1; number <= MAX_STATIONS; number++) list.append(settings().value(ipKey(number)).toString());
  return list;
}

QList<int> Settings::cvsInText(const QString &text)
{
  QList<int> cvs;
  const QStringList parts = text.split(QRegularExpression("[^0-9]+"), Qt::SkipEmptyParts);
  for (const QString &part : parts) {
    int cv = part.toInt();
    if ((cv >= 1) && (cv <= MAX_CV_NUMBER) && !cvs.contains(cv)) cvs.append(cv);
  }
  return cvs;
}

QList<int> Settings::chosenCvs()
{
  QList<int> cvs;
  for (int i = 0; i < numberOfCommonCvs; i++) {
    if (settings().value(checkedKey(commonCvs[i].cv), false).toBool()) cvs.append(commonCvs[i].cv);
  }
  for (int cv : cvsInText(settings().value("cvOther").toString())) if (!cvs.contains(cv)) cvs.append(cv);
  return cvs;
}

// The PoM address of the lift decoder follows from its first switch address: (switch address - 1) / 4 + 7000
int Settings::liftPomAddress()
{
  if (!settings().value("liftEnabled", false).toBool()) return 0;
  int switchAddress = settings().value("liftSwitchAddress").toString().toInt();
  if ((switchAddress < 1) || (switchAddress > 2048)) return 0;
  return (switchAddress - 1) / 4 + 7000;
}

int Settings::liftIrAddress()
{
  int address = settings().value("liftIrRsAddress").toString().toInt();
  if ((address < 1) || (address > 126)) return 0;
  return address;
}
