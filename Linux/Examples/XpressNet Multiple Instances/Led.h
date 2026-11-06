//
//  Led.h
//  XpressNetQt examples
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  A LED, as the Shape control in the VB6 examples and the Led control (Led.vb) in the .NET examples.
//  - fillColor: the colour of the LED (VB6: FillColor of the Shape control)
//  - round = true: a circle (VB6: Shape = 3), round = false: a rectangle (VB6: Shape = 0)
//  Around a round LED, the background of the parent (for example a QGroupBox) is visible.
//
//  The LED can be placed in a .ui file (Qt Designer: a QWidget, promoted to Led with header Led.h), or created in the
//  code, for example for a grid of many LEDs (such as the 128 x 8 feedback bits of the RS-Bus):
//    Led *led = new Led(parent);
//    led->setGeometry(x, y, 9, 9);
//    led->setRound(false);
//    led->setFillColor(Qt::green);
//

#ifndef LED_H
#define LED_H

#include <QColor>
#include <QWidget>

class Led : public QWidget
{
  Q_OBJECT
  Q_PROPERTY(bool round READ round WRITE setRound)
  Q_PROPERTY(QColor fillColor READ fillColor WRITE setFillColor)

public:
  explicit Led(QWidget *parent = nullptr);

  bool round() const;
  void setRound(bool value);                              // Circle or rectangle
  QColor fillColor() const;
  void setFillColor(const QColor &value);                 // A new colour is shown immediately
  QSize sizeHint() const override;                        // 25 x 25, as in .NET

protected:
  void paintEvent(QPaintEvent *event) override;

private:
  bool _round;
  QColor _fillColor;
};

#endif // LED_H
