//
//  Led.cpp
//  XpressNetQt examples
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  A LED, as the Shape control in the VB6 examples and the Led control (Led.vb) in the .NET examples. See Led.h
//

#include "Led.h"
#include <QPainter>


Led::Led(QWidget *parent) : QWidget(parent)
{
  _round = true;
  _fillColor = Qt::yellow;                                // As in .NET: a new LED is yellow
}

bool Led::round() const
{
  return _round;
}

void Led::setRound(bool value)
{
  _round = value;
  update();
}

QColor Led::fillColor() const
{
  return _fillColor;
}

void Led::setFillColor(const QColor &value)
{
  _fillColor = value;
  update();
}

QSize Led::sizeHint() const
{
  return QSize(25, 25);
}

// The widget itself has no background, so the background of the parent is visible around a round LED
void Led::paintEvent(QPaintEvent *event)
{
  Q_UNUSED(event);
  QPainter painter(this);
  painter.setPen(Qt::black);
  painter.setBrush(_fillColor);
  if (_round) {
    painter.setRenderHint(QPainter::Antialiasing);
    painter.drawEllipse(QRectF(0.5, 0.5, width() - 1, height() - 1));
  }
  else {
    painter.drawRect(0, 0, width() - 1, height() - 1);
  }
}
