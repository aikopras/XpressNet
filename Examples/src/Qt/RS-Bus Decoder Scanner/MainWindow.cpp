//
//  MainWindow.cpp
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of AppDelegate.m of the macOS version. The user interface:
//  - SCAN starts the scan of the command stations whose IP address is set in the preferences (1..5). During the
//    scan the button shows STOP, which ends the scan. The decoders found until then stay in the table.
//  - The progress bar and status line show how far the scan is.
//  - The table shows per decoder: the command station (IP address) that delivered the PoM messages, the PoM
//    address, the RS-Bus address(es), the type (CV27), the accessory addresses, and the CVs chosen in the
//    preferences. "?" means that the CV was not received. CmdStation (CV19) is shown by name.
//  - The columns can be moved and resized. Their order and widths are stored in the settings (columnOrder,
//    columnWidths), as is the size and position of the window (geometry).
//  - SAVE CSV saves the table as CSV file, with the date of the scan in the file name and on the first line.
//  The contents of the window are created in code (as in the macOS version), there is no .ui file.
//

#include "MainWindow.h"
#include "PreferencesDialog.h"
#include "Settings.h"
#include <QApplication>
#include <QCloseEvent>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QLabel>
#include <QMenuBar>
#include <QMessageBox>
#include <QProgressBar>
#include <QPushButton>
#include <QStandardPaths>
#include <QTableWidget>
#include <QTextBrowser>
#include <QTextStream>
#include <QVBoxLayout>

#define CV_COLUMN_WIDTH 90                  // Default width of a CV column

// The fixed columns of the table and the CSV file: identifier, name, second header line, default width
struct FixedColumn { const char *id; const char *name; const char *cv; int width; };
static const FixedColumn fixedColumns[] = {
  {"station", "Command station",     "",     130},
  {"pom",     "PoM address",         "",     90},
  {"rs",      "RS-Bus address",      "",     100},
  {"type",    "Type",                "CV27", 280},
  {"acc",     "Accessory addresses", "",     130},
};
static const int numberOfFixedColumns = sizeof(fixedColumns) / sizeof(fixedColumns[0]);

// The names of the generic CVs 1..32, as in CvValues.h of the AP_DCC_Decoder_Core library ("" = no name)
static const char *cvNames[] = {
  "",
  "myAddrL",      "",             "T_on_F1",      "T_on_F2",      "T_on_F3",      "T_on_F4",      "version",
  "VID",          "myAddrH",      "myRSAddr",     "DelayIn1",     "DelayIn2",     "DelayIn3",     "DelayIn4",
  "DelayIn5",     "DelayIn6",     "DelayIn7",     "DelayIn8",     "CmdStation",   "RSFEC",        "SkipUnEven",
  "RSParity",     "Search",       "RSPulsCount",  "Restart",      "DccQuality",   "DecType",      "RailCom",
  "Config",       "VID_2",        "ParityErrors", "PulseErrors",
};
static const int numberOfCvNames = sizeof(cvNames) / sizeof(cvNames[0]);


MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent)
{
  setWindowTitle("RS-Bus Decoder Scanner");
  Settings::checkDefaults();
  scanner = new DecoderScanner(this);
  connect(scanner, &DecoderScanner::scannerStatus, this, &MainWindow::scannerStatus);
  connect(scanner, &DecoderScanner::scannerFoundDecoder, this, &MainWindow::scannerFoundDecoder);
  connect(scanner, &DecoderScanner::scannerUpdatedDecoder, this, &MainWindow::scannerUpdatedDecoder);
  connect(scanner, &DecoderScanner::scannerFinished, this, &MainWindow::scannerFinished);
  createWindowContents();
  createColumns();
  // The size and position of the window
  if (!restoreGeometry(Settings::settings().value("geometry").toByteArray())) resize(1210, 600);
  labelStatus->setText("Push SCAN to start");
  updateButtons();
}

MainWindow::~MainWindow()
{
  delete helpWindow;
}

void MainWindow::closeEvent(QCloseEvent *event)
{
  scanner->stop();                                        // Closes the connections
  Settings::settings().setValue("geometry", saveGeometry());
  if (helpWindow != nullptr) helpWindow->close();
  event->accept();
}


// ************************************************************************************************************
// ********************************************* WINDOW CONTENTS **********************************************
// ************************************************************************************************************
void MainWindow::createWindowContents()
{
  // The menus
  QMenu *fileMenu = menuBar()->addMenu("&File");
  QAction *saveAction = fileMenu->addAction("Save CSV...", this, &MainWindow::saveCsv);
  saveAction->setShortcut(QKeySequence::Save);
  QAction *preferencesAction = fileMenu->addAction("Preferences...", this, &MainWindow::showPreferences);
  preferencesAction->setShortcut(QKeySequence::Preferences);
  preferencesAction->setMenuRole(QAction::PreferencesRole);
  fileMenu->addSeparator();
  QAction *quitAction = fileMenu->addAction("Quit", this, &QWidget::close);
  quitAction->setShortcut(QKeySequence::Quit);
  QMenu *helpMenu = menuBar()->addMenu("&Help");
  helpMenu->addAction("CV overview...", this, &MainWindow::showCvOverview);
  helpMenu->addAction("About RS-Bus Decoder Scanner", this, &MainWindow::showAbout);
  // SCAN / STOP, SAVE CSV and the progress bar (0..100)
  QWidget *central = new QWidget(this);
  QVBoxLayout *layout = new QVBoxLayout(central);
  QHBoxLayout *top = new QHBoxLayout;
  buttonScan = new QPushButton("SCAN");
  buttonScan->setMinimumWidth(110);
  connect(buttonScan, &QPushButton::clicked, this, &MainWindow::pushedScan);
  buttonSaveCsv = new QPushButton("SAVE CSV");
  buttonSaveCsv->setMinimumWidth(110);
  connect(buttonSaveCsv, &QPushButton::clicked, this, &MainWindow::saveCsv);
  progressBar = new QProgressBar;
  progressBar->setRange(0, 100);
  progressBar->setTextVisible(false);
  top->addWidget(buttonScan);
  top->addWidget(buttonSaveCsv);
  top->addSpacing(10);
  top->addWidget(progressBar, 1);
  layout->addLayout(top);
  // The status line
  labelStatus = new QLabel;
  labelStatus->setTextInteractionFlags(Qt::TextSelectableByMouse);
  labelStatus->setWordWrap(true);
  layout->addWidget(labelStatus);
  // The table. The columns are created by createColumns
  table = new QTableWidget(this);
  table->setEditTriggers(QAbstractItemView::NoEditTriggers);
  table->setAlternatingRowColors(true);
  table->setSelectionBehavior(QAbstractItemView::SelectRows);
  table->verticalHeader()->setVisible(false);
  table->horizontalHeader()->setSectionsMovable(true);
  table->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
  table->horizontalHeader()->setDefaultAlignment(Qt::AlignCenter);
  connect(table->horizontalHeader(), &QHeaderView::sectionMoved, this, [this]() { saveColumns(); });
  connect(table->horizontalHeader(), &QHeaderView::sectionResized, this, [this]() { saveColumns(); });
  layout->addWidget(table, 1);
  setCentralWidget(central);
}

void MainWindow::updateButtons()
{
  bool active = scanner->active();
  buttonScan->setText(active ? "STOP" : "SCAN");
  buttonSaveCsv->setEnabled(!active && !scanner->decoders().isEmpty());
}


// ************************************************************************************************************
// *********************************************** COLUMNS ****************************************************
// ************************************************************************************************************
// The header of a column: two lines (name and CV number), and one line in the CSV file
QString MainWindow::headerOfColumn(const QString &identifier, bool csv) const
{
  QString name;
  QString cvText;
  for (int i = 0; i < numberOfFixedColumns; i++) {
    if (identifier == fixedColumns[i].id) {
      name = fixedColumns[i].name;
      cvText = fixedColumns[i].cv;
    }
  }
  if (identifier.startsWith("cv")) {
    int cv = identifier.mid(2).toInt();
    if (cv < numberOfCvNames) name = cvNames[cv];
    cvText = QString("CV%1").arg(cv);
  }
  if (!csv) return name + "\n" + cvText;
  if (name.isEmpty()) return cvText;
  if (cvText.isEmpty()) return name;
  return QString("%1 (%2)").arg(name, cvText);
}

// Creates the columns: the fixed columns and the CVs of the preferences. The order and widths come from the
// settings (columnOrder, columnWidths); new columns are added at the end, with their default width
void MainWindow::createColumns()
{
  buildingColumns = true;
  QSettings &s = Settings::settings();
  QStringList savedOrder = s.value("columnOrder").toStringList();
  QVariantMap savedWidths = s.value("columnWidths").toMap();
  // All columns with their default widths, in the default order
  QStringList identifiers;
  QMap<QString, int> widths;
  for (int i = 0; i < numberOfFixedColumns; i++) {
    identifiers.append(fixedColumns[i].id);
    widths[fixedColumns[i].id] = fixedColumns[i].width;
  }
  for (int cv : Settings::chosenCvs()) {
    QString identifier = QString("cv%1").arg(cv);
    identifiers.append(identifier);
    widths[identifier] = CV_COLUMN_WIDTH;
  }
  // The order: first the stored order, then the new columns
  columnIds.clear();
  for (const QString &identifier : savedOrder) {
    if (identifiers.contains(identifier) && !columnIds.contains(identifier)) columnIds.append(identifier);
  }
  for (const QString &identifier : identifiers) if (!columnIds.contains(identifier)) columnIds.append(identifier);
  // Remove the old columns (this also resets the moved sections), and add the new ones
  table->setRowCount(0);
  table->setColumnCount(0);
  table->setColumnCount(columnIds.count());
  for (int i = 0; i < columnIds.count(); i++) {
    table->setHorizontalHeaderItem(i, new QTableWidgetItem(headerOfColumn(columnIds[i], false)));
    int width = savedWidths.contains(columnIds[i]) ? savedWidths[columnIds[i]].toInt() : widths[columnIds[i]];
    table->setColumnWidth(i, width);
  }
  buildingColumns = false;
}

// Stores the order and widths of the columns, after the user moved or resized a column
void MainWindow::saveColumns()
{
  if (buildingColumns) return;
  QSettings &s = Settings::settings();
  QStringList order;
  QVariantMap widths = s.value("columnWidths").toMap();
  QHeaderView *header = table->horizontalHeader();
  for (int visual = 0; visual < header->count(); visual++) {
    int logical = header->logicalIndex(visual);
    order.append(columnIds[logical]);
    widths[columnIds[logical]] = header->sectionSize(logical);
  }
  s.setValue("columnOrder", order);
  s.setValue("columnWidths", widths);
}


// ************************************************************************************************************
// ************************************************ TABLE *****************************************************
// ************************************************************************************************************
// The text of a column, for the table and the CSV file
QString MainWindow::textForColumn(const QString &identifier, const ScannedDecoder *decoder) const
{
  // The IR sensors of the lift: their own RS-Bus addresses and type, the other columns of the lift decoder
  if (decoder->liftIr && (identifier != "rs") && (identifier != "type")) decoder = decoder->liftLevel;
  if (identifier == "station") return decoder->pomIpAddress;
  if (identifier == "pom")     return QString::number(decoder->pomAddress());
  if (identifier == "rs")      return decoder->rsAddressText();
  if (identifier == "type")    return decoder->typeText();
  if (identifier == "acc")     return decoder->accessoryAddressText();
  if (identifier == "cv19")    return decoder->cmdStationText();
  if (identifier.startsWith("cv")) return decoder->cvText(identifier.mid(2).toInt());
  return QString();
}

void MainWindow::showRow(int row)
{
  const ScannedDecoder *decoder = scanner->decoders()[row];
  for (int column = 0; column < columnIds.count(); column++) {
    QTableWidgetItem *item = table->item(row, column);
    if (item == nullptr) {
      item = new QTableWidgetItem;
      item->setTextAlignment(Qt::AlignCenter);
      table->setItem(row, column, item);
    }
    item->setText(textForColumn(columnIds[column], decoder));
  }
}

void MainWindow::showTable()
{
  table->setRowCount(scanner->decoders().count());
  for (int row = 0; row < scanner->decoders().count(); row++) showRow(row);
}


// ************************************************************************************************************
// ************************************************ SCAN ******************************************************
// ************************************************************************************************************
void MainWindow::pushedScan()
{
  if (scanner->active()) {
    scanner->stop();
    return;
  }
  // The settings are read each time, so new values are used after SCAN is pushed. First the new columns, then the
  // scan: the scan deletes the decoders of the previous scan
  createColumns();
  scanDate = QDateTime::currentDateTime();
  progressBar->setValue(0);
  scanner->start(Settings::ipAddresses(), Settings::chosenCvs(), Settings::liftPomAddress(), Settings::liftIrAddress());
  showTable();
  updateButtons();
}

void MainWindow::scannerStatus(QString text, double progress)
{
  labelStatus->setText(text);
  progressBar->setValue(qRound(progress));
}

void MainWindow::scannerFoundDecoder(ScannedDecoder *)
{
  showTable();
}

void MainWindow::scannerUpdatedDecoder(ScannedDecoder *decoder)
{
  int row = scanner->decoders().indexOf(decoder);
  if (row >= 0) showRow(row);
}

void MainWindow::scannerFinished(QString text)
{
  labelStatus->setText(text);
  progressBar->setValue(100);
  showTable();                                            // The table has been sorted
  updateButtons();
}


// ************************************************************************************************************
// ************************************************ CSV *******************************************************
// ************************************************************************************************************
// A field between quotes, if needed (RFC 4180)
static QString csvField(const QString &text)
{
  if (!text.contains(',') && !text.contains('"') && !text.contains('\n')) return text;
  QString quoted = text;
  quoted.replace("\"", "\"\"");
  return "\"" + quoted + "\"";
}

void MainWindow::saveCsv()
{
  if (scanner->active() || scanner->decoders().isEmpty()) return;
  QString directory = QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation);
  QString name = QString("RS-Bus Decoders %1.csv").arg(scanDate.toString("yyyy-MM-dd"));
  QString path = QFileDialog::getSaveFileName(this, "Save CSV", directory + "/" + name, "CSV files (*.csv)");
  if (path.isEmpty()) return;
  // First line: the date of the scan. Second line: the column titles. Then one line per decoder.
  // The columns are in the order of the table
  QHeaderView *header = table->horizontalHeader();
  QStringList lines;
  lines.append(QString("RS-Bus Decoder Scanner,Scan of %1").arg(scanDate.toString("yyyy-MM-dd HH:mm")));
  QStringList fields;
  for (int visual = 0; visual < header->count(); visual++) fields.append(csvField(headerOfColumn(columnIds[header->logicalIndex(visual)], true)));
  lines.append(fields.join(","));
  for (const ScannedDecoder *decoder : scanner->decoders()) {
    fields.clear();
    for (int visual = 0; visual < header->count(); visual++) fields.append(csvField(textForColumn(columnIds[header->logicalIndex(visual)], decoder)));
    lines.append(fields.join(","));
  }
  QFile file(path);
  if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
    labelStatus->setText("CSV file not saved: " + file.errorString());
    return;
  }
  QTextStream out(&file);
  out << lines.join("\n") << "\n";
  labelStatus->setText("Saved: " + QDir::toNativeSeparators(path));
}


// ************************************************************************************************************
// ******************************************* PREFERENCES AND HELP *******************************************
// ************************************************************************************************************
void MainWindow::showPreferences()
{
  PreferencesDialog dialog(this);
  connect(&dialog, &PreferencesDialog::cvOverviewRequested, this, &MainWindow::showCvOverview);
  dialog.exec();
}

// The CV overview: the same HTML file as in the macOS version (a resource, see CMakeLists.txt)
void MainWindow::showCvOverview()
{
  if (helpWindow == nullptr) {
    helpWindow = new QTextBrowser;
    helpWindow->setWindowTitle("RS-Bus Decoder Scanner - CV overview");
    helpWindow->setSource(QUrl("qrc:/help/CV-Overview.html"));
    helpWindow->resize(900, 640);
  }
  helpWindow->show();
  helpWindow->raise();
  helpWindow->activateWindow();
}

void MainWindow::showAbout()
{
  QMessageBox::about(this, "About RS-Bus Decoder Scanner",
                     QString("<b>RS-Bus Decoder Scanner %1</b><br><br>Finds all decoders with RS-Bus feedback on up to 5 "
                             "command stations (Lenz LAN/USB interfaces), and reads their type, software version and "
                             "some other CVs via PoM. The results can be saved as CSV file.<br><br>"
                             "Example program of the XpressNet library: "
                             "<a href=\"https://github.com/aikopras/XpressNet\">github.com/aikopras/XpressNet</a>"
                             "<br><br>Copyright (c) 2026 by Aiko Pras").arg(APP_VERSION));
}
