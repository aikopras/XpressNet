# XpressNetQt example XpressNet Basic Test (version 2.0, November 2026)
# Build (see Examples.md):  cmake -S . -B build  &&  cmake --build build  &&  "./build/XpressNet Basic Test"
# Requires Qt 6 (Widgets, Network, SerialPort) and a C++17 compiler.

cmake_minimum_required(VERSION 3.16)
project(BasicTest VERSION 2.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_INCLUDE_CURRENT_DIR ON)                 # The generated ui_*.h includes Led.h

find_package(Qt6 REQUIRED COMPONENTS Widgets)

# The library. Not needed if this example is built from Linux/CMakeLists.txt (the library is then built there)
if(NOT TARGET XpressNetQt)
  add_subdirectory(../../Code/XpressNetQt XpressNetQt)
endif()

add_executable(BasicTest
  main.cpp
  FrmBasicTest.h
  FrmBasicTest.cpp
  FrmBasicTest.ui
  Led.h
  Led.cpp
)

target_link_libraries(BasicTest PRIVATE Qt6::Widgets XpressNetQt)
# The name of the program, the same as in the other versions (VB6, .NET, macOS)
set_target_properties(BasicTest PROPERTIES OUTPUT_NAME "XpressNet Basic Test")
