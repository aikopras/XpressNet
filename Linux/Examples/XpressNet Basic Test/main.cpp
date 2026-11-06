//
//  main.cpp
//  XpressNetQt examples: XpressNet Basic Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Starts the form FrmBasicTest (.NET: the startup form of the project)
//

#include <QApplication>
#include "FrmBasicTest.h"

int main(int argc, char *argv[])
{
  QApplication app(argc, argv);
  FrmBasicTest form;
  form.show();
  return app.exec();
}
