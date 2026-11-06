//
//  PanelManager.cpp
//  XpressNetQt examples: RS-Bus Monitor
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of the class PanelManager of the RS-Bus Monitor (VB6: PanelManager.cls, .NET: PanelManager.vb).
//  The layout, the margins and the colours are the same as in VB6 and .NET. Differences with .NET:
//  - picPanel is a QWidget. Its Paint event is caught by eventFilter (.NET: Handles m_PicPanel.Paint), in which the
//    whole panel is drawn with a QPainter (DrawAll). A change of a LED or of a square asks for a new drawing (update).
//  - The text is measured with QFontMetrics (.NET: TextRenderer).
//  - The arrays with the centers have a fixed size (VB6 / .NET: ReDim in PrecomputeCenters).
//

#include "PanelManager.h"
#include <QEvent>
#include <QFontMetrics>
#include <QPainter>
#include <QWidget>
#include <cmath>

PanelManager::PanelManager(QObject *parent) : QObject(parent)
{
  // VB6 / .NET: the variables start with 0
  m_RowH = 0; m_D = 0; m_R = 0; m_Spacing = 0; m_BlockGap = 0; m_LabelW = 0; m_BlockW = 0; m_BlockPad = 0; m_Sq = 0;
  for (int a = 0; a <= 127; a++) {
    m_Bx[a] = 0; m_By[a] = 0;
    for (int k = 0; k <= 7; k++) { m_Cx[a][k] = 0; m_Cy[a][k] = 0; }
  }
  m_PicPanel = nullptr;
}

// ======= Public Methods =======

// Initialiseert de PanelManager met een QWidget (VB6 / .NET: een PictureBox).
// @param picPanel: De QWidget waarin getekend wordt.
void PanelManager::Init(QWidget *picPanel)
{
  m_PicPanel = picPanel;
  m_PicPanel->installEventFilter(this);
  InitializeColors();
  InitializeFont();
  InitializeState();
  RebuildLayoutAndRedraw();
}


// Event welke de 4 LEDs toont die bij een bepaalde feedback nibble behoren
// @param address: 1..128
// @param nibble: 0 / 1 (0 = data is for bits 0..3 / 1 = data is for bits 4..7)
// @param data: 0..15
void PanelManager::SetFeedbackNibble(int address, quint8 nibble, quint8 data, BlockStateType bState)
{
  if (address < 1 || address > 128) return;
  if (nibble > 1) return;
  if (data > 15) return;
  address = address - 1;                   // Internally addresses are from 0..127
  // Ik moet alle vier de bits afzonderlijk testen en een boolean zetten als er iets wordt aangepast
  // Als Nibble = 0, dan heb ik bits 0..3. Als Nibble = 1, dan heb ik bits 4..7
  bool refreshNeeded = false;
  LedStateType value;                      // becomes 0 (LED_OFF) or 1 (LED_ON)
  quint8 offset;                           // becomes 0 or 4
  if (nibble == 0) offset = 0; else offset = 4;
  for (quint8 i = 0; i <= 3; i++) {
    value = static_cast<LedStateType>(bitValue(data, i));
    if (LedState[address][i + offset] != value) {
      LedState[address][i + offset] = value;
      refreshNeeded = true;
    }
  }
  // Kleur nu het blokje voor de LEDs
  if (blockState[address] != bState) {
    blockState[address] = bState;
    refreshNeeded = true;
  }
  if (refreshNeeded) m_PicPanel->update();
}

//---------------------------------------------------------------------------------------------------------------------
// bit manipulation function: returns 0 or 1
quint8 PanelManager::bitValue(quint8 data, quint8 index)
{
  return (data >> index) & 1;
}

// Stelt de staat van een LED in.
// @param addr: Adres (1..128).
// @param bit: Bit (0..7).
// @param state: Nieuwe staat (LED_OFF, LED_ON, LED_NOT_CONNECTED).
void PanelManager::SetLed(int addr, int bit, LedStateType state)
{
  addr = addr - 1;
  if (addr < 0 || addr > 127) return;
  if (bit < 0 || bit > 7) return;
  if (LedState[addr][bit] != state) {
    LedState[addr][bit] = state;
    m_PicPanel->update();
  }
}

// Stelt de staat van een Blokje in.
// @param addr: Adres (1..128).
// @param state: Nieuwe staat (NOT_CONNECTED, CENTRALE, WISSEL, BEZETMELDING).
void PanelManager::SetBlock(int addr, BlockStateType state)
{
  addr = addr - 1;
  if (addr < 0 || addr > 127) return;
  if (blockState[addr] != state) {
    blockState[addr] = state;
    m_PicPanel->update();
  }
}

// Hertekent de layout en inhoud van picPanel.
void PanelManager::Resize()
{
  RebuildLayoutAndRedraw();
}

// ======= Private Methods =======

// Initialiseert de kleuren.
void PanelManager::InitializeColors()
{
  COLOR_BACKGROUND = QColor(240, 240, 240);
  COLOR_LABEL = QColor(0, 0, 255);
  COLOR_LED_ON = QColor(220, 30, 30);
  COLOR_LED_OFF = QColor(192, 242, 154);
  COLOR_LED_NC = QColor(220, 220, 220);
  COLOR_LED_EDGE = QColor(100, 100, 100);
  COLOR_LED_GLOW = QColor(255, 200, 200);
  // TODO: Betere kleuren geven
  COLOR_BLOCK_NC = QColor(220, 220, 220);
  COLOR_CENTALE_NC = QColor(255, 0, 0);
  COLOR_WISSEL_NC = QColor(0, 255, 0);
  COLOR_BEZETMELDING = QColor(0, 0, 255);
}

// Initialiseert het lettertype voor labels.
// Linux: Arial (Liberation Sans), as in the .ui files of the examples (VB6: MS Sans Serif)
void PanelManager::InitializeFont()
{
  m_Font = QFont("Arial", 10);
  m_Font.setBold(true);
}

// Initialiseert de beginstate.
void PanelManager::InitializeState()
{
  for (int a = 0; a <= 127; a++) {
    blockState[a] = BlockStateType::NOT_CONNECTED;
    for (int k = 0; k <= 7; k++) {
      LedState[a][k] = LedStateType::LED_NOT_CONNECTED;
    }
  }
}

// Herberekent de layout en tekent alles opnieuw.
void PanelManager::RebuildLayoutAndRedraw()
{
  ComputeLayout(m_PicPanel->width(), m_PicPanel->height());
  PrecomputeCenters();
  m_PicPanel->update();
}

// Bereken de layout van het paneel.
// @param W: Breedte van picPanel.
// @param H: Hoogte van picPanel.
void PanelManager::ComputeLayout(double W, double H)
{
  double contentW, contentH;
  contentW = W - (MARGIN_L + MARGIN_R);
  contentH = H - (MARGIN_T + MARGIN_B);
  if (contentW < 100 || contentH < 100) return;

  // Bereken de breedte voor de labelkolom
  m_LabelW = QFontMetrics(m_Font).horizontalAdvance("128") + LABEL_OFFSET + 6;

  // Bereken de hoogte per rij
  m_RowH = CalculateRowHeight(contentH);

  // Bereken de diameter van de bits en de spacing
  m_D = CalculateBitDiameter(m_RowH);
  m_R = m_D / 2;
  m_Spacing = CalculateSpacing(m_D);

  // Bereken de grootte van de blokjes
  m_Sq = CalculateBlockSize(m_RowH, m_R);

  // Bereken de breedte van een blok
  CalculateBlockWidth(contentW);
}

// Bereken de hoogte per rij.
double PanelManager::CalculateRowHeight(double contentH)
{
  double rowH = std::floor(contentH / ROWS_PER_BLOCK);    // VB: Int
  if (rowH < 8) rowH = 8;
  return rowH;
}

// Bereken de diameter van de bits.
double PanelManager::CalculateBitDiameter(double rowH)
{
  double bitDiameter = rowH - 4;
  if (bitDiameter < 6) bitDiameter = 6;
  return bitDiameter;
}

// Bereken de spacing tussen bits.
double PanelManager::CalculateSpacing(double bitDiameter)
{
  double spacing = std::round(bitDiameter / 3);
  if (spacing < 2) spacing = 2;
  return spacing;
}

// Bereken de grootte van de blokjes.
double PanelManager::CalculateBlockSize(double rowH, double radius)
{
  double blockSize = rowH / 2;
  if (blockSize < 2) blockSize = 2;
  if (blockSize > (radius / 3)) blockSize = radius / 3;
  return blockSize;
}

// Bereken de breedte van een blok.
void PanelManager::CalculateBlockWidth(double contentW)
{
  m_BlockGap = 24;
  m_BlockPad = m_Spacing / 4;
  if (m_BlockPad < 1) m_BlockPad = 1;

  // Iteratief de breedte van een blok berekenen
  while (true) {
    m_BlockW = m_LabelW + (BITS_PER_ADDR * m_D) +
               ((BITS_PER_ADDR - 1) * m_Spacing) + MIDGROUP_GAP + m_BlockPad + m_Sq;
    double totalW;
    totalW = (NUM_BLOCKS * m_BlockW) + ((NUM_BLOCKS - 1) * m_BlockGap);
    if (totalW <= contentW) break;
    // Pas de afmetingen aan als het niet past
    if (m_D > 6) m_D = m_D - 1;
    m_R = m_D / 2;
    m_Spacing = CalculateSpacing(m_D);
    if (m_BlockGap > 12) m_BlockGap = m_BlockGap - 1;
    if (m_D <= 6 && m_BlockGap <= 12) break;
  }
}

// Bepaal de exacte pixelcoördinaten voor LEDs en blokjes.
void PanelManager::PrecomputeCenters()
{
  double baseX, baseY;
  double ledOffsetX;
  ledOffsetX = m_BlockPad + m_Sq * 3;

  for (int b = 0; b <= NUM_BLOCKS - 1; b++) {
    baseX = MARGIN_L + b * (m_BlockW + m_BlockGap);
    for (int r = 0; r <= ROWS_PER_BLOCK - 1; r++) {
      baseY = MARGIN_T + r * m_RowH;
      int addr = b * ROWS_PER_BLOCK + r;

      // Blokjes voor terugmeld type
      m_Bx[addr] = baseX + m_LabelW + m_BlockPad + m_Sq / 2;
      m_By[addr] = baseY + m_RowH / 2;

      // LEDs
      for (int k = 0; k <= BITS_PER_ADDR - 1; k++) {
        double x, y;
        x = baseX + m_LabelW + ledOffsetX + k * (m_D + m_Spacing) + m_R;
        if (k >= 4) x = x + MIDGROUP_GAP;
        y = baseY + (m_RowH / 2);
        m_Cx[addr][k] = x;
        m_Cy[addr][k] = y;
      }
    }
  }
}

// .NET: m_PicPanel_Paint. The whole panel is drawn after each change
bool PanelManager::eventFilter(QObject *watched, QEvent *event)
{
  if ((watched == m_PicPanel) && (event->type() == QEvent::Paint)) {
    QPainter p(m_PicPanel);
    DrawAll(p);
    return true;
  }
  return QObject::eventFilter(watched, event);
}

// Tekent alles in picPanel.
void PanelManager::DrawAll(QPainter &p)
{
  // Wis de achtergrond
  p.fillRect(m_PicPanel->rect(), COLOR_BACKGROUND);
  p.setRenderHint(QPainter::Antialiasing);
  p.setFont(m_Font);
  QFontMetrics fm(m_Font);

  // Teken alle adressen, labels, blokjes en LEDs
  QString labelText;
  double labelX, labelY;
  int b, r;

  for (int addr = 0; addr <= (NUM_BLOCKS * ROWS_PER_BLOCK) - 1; addr++) {
    b = addr / ROWS_PER_BLOCK;
    r = addr % ROWS_PER_BLOCK;

    // Teken labels
    if ((r % 4) == 0) {
      labelX = (MARGIN_L + b * (m_BlockW + m_BlockGap)) + LABEL_OFFSET;
      labelY = MARGIN_T + r * m_RowH + (int(m_RowH) / 2) - (fm.height() / 2);
      labelText = QString::number(addr + 1);
      p.setPen(COLOR_LABEL);
      p.drawText(QPointF(labelX, labelY + fm.ascent()), labelText);   // drawText: the position of the baseline
    }

    // Teken blokjes
    DrawBlock(p, addr, blockState[addr]);

    // Teken LEDs
    for (int k = 0; k <= BITS_PER_ADDR - 1; k++) {
      DrawLamp(p, m_Cx[addr][k], m_Cy[addr][k], LedState[addr][k]);
    }
  }
}

// Tekent een blokje voor een adres.
// @param addr: Het adres (0..127).
void PanelManager::DrawBlock(QPainter &p, int addr, BlockStateType state)
{
  if (addr < 0 || addr > 127) return;
  QColor fillC;
  switch (state) {
    case BlockStateType::NOT_CONNECTED: fillC = COLOR_BLOCK_NC; break;
    case BlockStateType::CENTRALE: fillC = COLOR_CENTALE_NC; break;
    case BlockStateType::WISSEL: fillC = COLOR_WISSEL_NC; break;
    default: fillC = COLOR_BEZETMELDING; break;
  }
  double bx, by, s;
  bx = m_Bx[addr];
  by = m_By[addr];
  s = m_Sq;
  p.fillRect(QRectF(bx - s / 2, by - s / 2, s, s), fillC);
}

// Tekent een LED-lampje.
// @param cx: X-coördinaat van het midden.
// @param cy: Y-coördinaat van het midden.
// @param state: Staat van de LED (LED_OFF, LED_ON, LED_NOT_CONNECTED).
void PanelManager::DrawLamp(QPainter &p, double cx, double cy, LedStateType state)
{
  QColor fillC;
  switch (state) {
    case LedStateType::LED_ON: fillC = COLOR_LED_ON; break;
    case LedStateType::LED_NOT_CONNECTED: fillC = COLOR_LED_NC; break;
    default: fillC = COLOR_LED_OFF; break;
  }

  Circle(p, cx, cy, m_R, fillC, COLOR_LED_EDGE);

  // Teken de glans als de LED aan staat en een subtiele glans toe als de LED uit is
  if (state == LedStateType::LED_ON && m_R >= 6) {
    Circle(p, cx - m_R / 3, cy - m_R / 3, m_R / 3, COLOR_LED_GLOW, fillC);
  } else if (state == LedStateType::LED_OFF && m_R >= 6) {
    Circle(p, cx - m_R / 3, cy - m_R / 3, m_R / 4, QColor(200, 200, 200), fillC);   // Kleinere glans
  }
}

// VB6: FillColor, FillStyle = vbFSSolid and Circle (cx, cy), radius, edgeColor
void PanelManager::Circle(QPainter &p, double cx, double cy, double radius, const QColor &fillC, const QColor &edgeC)
{
  p.setBrush(fillC);
  p.setPen(edgeC);
  p.drawEllipse(QPointF(cx, cy), radius, radius);
}
