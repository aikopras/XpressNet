//
//  main.cpp
//  XpressNetQt examples: Z21 Basic Test
//
//  Created by Aiko Pras on 05-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Starts the form FrmZ21Test (.NET: the startup form of the project)
//

#include <QApplication>
#include "FrmZ21Test.h"

int main(int argc, char *argv[])
{
  QApplication app(argc, argv);
  FrmZ21Test form;
  form.show();
  return app.exec();
}
