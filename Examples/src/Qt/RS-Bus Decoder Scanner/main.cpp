//
//  main.cpp
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Starts the main window (macOS: MainMenu.xib and AppDelegate)
//

#include <QApplication>
#include <QIcon>
#include "MainWindow.h"

int main(int argc, char *argv[])
{
  QApplication app(argc, argv);
  QApplication::setApplicationName("RS-Bus Decoder Scanner");
  QApplication::setApplicationVersion(APP_VERSION);
  QApplication::setWindowIcon(QIcon(":/RS-Bus-Decoder-Scanner.png"));
  MainWindow window;
  window.show();
  return app.exec();
}
