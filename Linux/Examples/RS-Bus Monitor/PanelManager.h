//
//  PanelManager.h
//  XpressNetQt examples: RS-Bus Monitor
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See PanelManager.cpp
//

#ifndef PANELMANAGER_H
#define PANELMANAGER_H

#include <QColor>
#include <QFont>
#include <QObject>

class QPainter;
class QWidget;

// ======= Enums voor betere leesbaarheid =======
enum class LedStateType {
  LED_OFF = 0,
  LED_ON = 1,
  LED_NOT_CONNECTED = 2
};

enum class BlockStateType {
  NOT_CONNECTED = 0,
  CENTRALE = 1,
  WISSEL = 2,
  BEZETMELDING = 3
};

class PanelManager : public QObject
{
  Q_OBJECT

public:
  explicit PanelManager(QObject *parent = nullptr);

  void Init(QWidget *picPanel);
  void SetFeedbackNibble(int address, quint8 nibble, quint8 data, BlockStateType bState);
  void SetLed(int addr, int bit, LedStateType state);
  void SetBlock(int addr, BlockStateType state);
  void Resize();

protected:
  bool eventFilter(QObject *watched, QEvent *event) override;  // The Paint event of picPanel

private:
  // ======= Configuratie voor de inhoud van picPanel =======
  static const int NUM_BLOCKS = 4;
  static const int ROWS_PER_BLOCK = 32;
  static const int BITS_PER_ADDR = 8;
  // Marges in pixels
  static const int MARGIN_L = 12;
  static const int MARGIN_T = 12;
  static const int MARGIN_R = 12;
  static const int MARGIN_B = 12;
  // Extra offsets (specifiek voor de inhoud van picPanel)
  static const int LABEL_OFFSET = 8;
  static const int MIDGROUP_GAP = 6;

  // ======= Stijl / Kleuren =======
  QColor COLOR_BACKGROUND;
  QColor COLOR_LABEL;
  QColor COLOR_LED_ON;
  QColor COLOR_LED_OFF;
  QColor COLOR_LED_NC;
  QColor COLOR_LED_EDGE;
  QColor COLOR_LED_GLOW;
  QColor COLOR_BLOCK_NC;
  QColor COLOR_CENTALE_NC;
  QColor COLOR_WISSEL_NC;
  QColor COLOR_BEZETMELDING;

  // ======= Layout-variabelen =======
  double m_RowH;
  double m_D;
  double m_R;
  double m_Spacing;
  double m_BlockGap;
  double m_LabelW;
  double m_BlockW;
  double m_BlockPad;
  double m_Sq;

  // ======= State =======
  LedStateType LedState[128][8];
  BlockStateType blockState[128];

  // ======= Centers voor LEDs en blokjes =======
  double m_Cx[NUM_BLOCKS * ROWS_PER_BLOCK][BITS_PER_ADDR];
  double m_Cy[NUM_BLOCKS * ROWS_PER_BLOCK][BITS_PER_ADDR];
  double m_Bx[128];
  double m_By[128];

  // ======= Referentie naar picPanel =======
  QWidget *m_PicPanel;
  QFont m_Font;                                           // VB6: the font of picPanel

  static quint8 bitValue(quint8 data, quint8 index);
  void InitializeColors();
  void InitializeFont();
  void InitializeState();
  void RebuildLayoutAndRedraw();
  void ComputeLayout(double W, double H);
  double CalculateRowHeight(double contentH);
  double CalculateBitDiameter(double rowH);
  double CalculateSpacing(double bitDiameter);
  double CalculateBlockSize(double rowH, double radius);
  void CalculateBlockWidth(double contentW);
  void PrecomputeCenters();
  void DrawAll(QPainter &p);
  void DrawBlock(QPainter &p, int addr, BlockStateType state);
  void DrawLamp(QPainter &p, double cx, double cy, LedStateType state);
  void Circle(QPainter &p, double cx, double cy, double radius, const QColor &fillC, const QColor &edgeC);
};

#endif // PANELMANAGER_H
