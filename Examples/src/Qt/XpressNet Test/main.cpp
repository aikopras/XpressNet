//
//  main.cpp
//  XpressNetQt examples: XpressNet Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Starts the form FrmExtendedTest (.NET: the startup form of the project)
//

#include <QApplication>
#include "FrmExtendedTest.h"

int main(int argc, char *argv[])
{
  QApplication app(argc, argv);
  FrmExtendedTest form;
  form.show();
  return app.exec();
}
