//
//  main.cpp
//  XpressNetQt examples: XpressNet Multiple Instances
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Starts the form FrmStart (.NET: the startup form of the project)
//

#include <QApplication>
#include "FrmStart.h"

int main(int argc, char *argv[])
{
  QApplication app(argc, argv);
  FrmStart form;
  form.show();
  return app.exec();
}
