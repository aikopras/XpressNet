//
//  Settings.h
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The preferences, with QSettings (Windows: registry HKEY_CURRENT_USER\Software\RS-Bus Decoder Scanner\Settings,
//  Linux: ~/.config/RS-Bus Decoder Scanner/Settings.conf). The keys are the same as in the macOS version:
//  - ipAddressStation1 .. ipAddressStation5: the IP addresses of the command stations (empty = not used)
//  - cvChecked<n>: the common CVs that are shown in the table; cvOther: the other CVs, as numbers ("33, 34")
//  - liftEnabled, liftSwitchAddress, liftIrRsAddress: the lift decoder
//  - columnOrder, columnWidths: the order and widths of the columns; geometry: the size and position of the window
//

#ifndef SETTINGS_H
#define SETTINGS_H

#include <QSettings>
#include <QStringList>
#include <QList>

// The common CVs, that can be chosen with a checkbox in the preferences, in the order of the preferences window.
// The other CVs can be entered as numbers. CV23 (Search) and CV25 (Restart) are not stored in EEPROM, and therefore
// not offered. CV1, CV9, CV10 and CV27 are shown in the fixed columns. * = version 20 or higher only
struct CommonCv { int cv; const char *name; bool newOnly; };
static const CommonCv commonCvs[] = {
  {7, "version", false}, {8, "VID", false}, {19, "CmdStation", false}, {20, "RSFEC", false},
  {21, "SkipUnEven", false}, {22, "RSParity", true}, {24, "RSPulsCount", true}, {26, "DccQuality", false},
  {28, "RailCom", false}, {29, "Config", false}, {30, "VID_2", false}, {31, "ParityErrors", true},
  {32, "PulseErrors", true},
};
static const int numberOfCommonCvs = sizeof(commonCvs) / sizeof(commonCvs[0]);
static const int defaultCvs[] = {7, 19, 20, 26};

class Settings
{
public:
  static QSettings &settings();
  static void checkDefaults();                            // Each setting that does not exist yet gets its default value
  static QString ipKey(int number) { return QString("ipAddressStation%1").arg(number); }
  static QString checkedKey(int cv) { return QString("cvChecked%1").arg(cv); }
  static QStringList ipAddresses();                       // The IP addresses of the command stations 1..5
  static QList<int> cvsInText(const QString &text);       // The CV numbers in a text, in the given order, once each
  static QList<int> chosenCvs();                          // The checked common CVs, then the other CVs
  static int liftPomAddress();                            // 0 = the lift decoder is not included
  static int liftIrAddress();                             // 0 = no IR sensors
};

#endif // SETTINGS_H
