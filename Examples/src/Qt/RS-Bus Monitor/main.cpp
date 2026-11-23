//
//  main.cpp
//  XpressNetQt examples: RS-Bus Monitor
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Starts the form FrmRsBusMonitor (.NET: the startup form of the project)
//

#include <QApplication>
#include "FrmRsBusMonitor.h"

int main(int argc, char *argv[])
{
  QApplication app(argc, argv);
  FrmRsBusMonitor form;
  form.show();
  return app.exec();
}
