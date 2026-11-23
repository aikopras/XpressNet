//
//  MainWindow.h
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See MainWindow.cpp
//

#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDateTime>
#include "DecoderScanner.h"

class QPushButton;
class QProgressBar;
class QLabel;
class QTableWidget;
class QTextBrowser;

class MainWindow : public QMainWindow
{
  Q_OBJECT

public:
  explicit MainWindow(QWidget *parent = nullptr);
  ~MainWindow() override;

protected:
  void closeEvent(QCloseEvent *event) override;

private slots:
  void pushedScan();                                      // SCAN, or STOP during a scan
  void saveCsv();
  void showPreferences();
  void showCvOverview();
  void showAbout();
  // Signals of the scanner
  void scannerStatus(QString text, double progress);
  void scannerFoundDecoder(ScannedDecoder *decoder);
  void scannerUpdatedDecoder(ScannedDecoder *decoder);
  void scannerFinished(QString text);

private:
  DecoderScanner *scanner;
  QDateTime scanDate;
  QPushButton *buttonScan;
  QPushButton *buttonSaveCsv;
  QProgressBar *progressBar;
  QLabel *labelStatus;
  QTableWidget *table;
  QTextBrowser *helpWindow = nullptr;                     // The CV overview
  QStringList columnIds;                                  // The identifiers of the columns, in logical order
  bool buildingColumns = false;                           // No saving of column order and widths meanwhile

  void createWindowContents();
  void createColumns();
  void saveColumns();
  void updateButtons();
  void showRow(int row);
  void showTable();
  QString headerOfColumn(const QString &identifier, bool csv) const;
  QString textForColumn(const QString &identifier, const ScannedDecoder *decoder) const;
};

#endif // MAINWINDOW_H
