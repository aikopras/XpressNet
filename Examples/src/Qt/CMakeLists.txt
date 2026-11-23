# XpressNet Qt (Linux, Windows): the library and the example programs (version 2.0, November 2026)
# Build (see Library/src/Qt/Install/Install-Library.md), in this directory:
#   cmake -S . -B build
#   cmake --build build
# The programs are then in build/<example>/, with the name of the example. Requires Qt 6 (Widgets, Network, SerialPort) and C++17.
# The library and every example can also be built separately, with the CMakeLists.txt in its own directory.

cmake_minimum_required(VERSION 3.16)
project(XpressNetLinux VERSION 2.0 LANGUAGES CXX)

add_subdirectory(../../../Library/src/Qt/XpressNetQt XpressNetQt)
add_subdirectory("XpressNet Basic Test")
add_subdirectory("XpressNet Test")
add_subdirectory("Z21 Basic Test")
add_subdirectory("XpressNet Multiple Instances")
add_subdirectory("RS-Bus Monitor")
add_subdirectory("RS-Bus Decoder Scanner")
