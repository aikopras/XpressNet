# XpressNet for Linux (Qt): the library and the five example programs (version 2.0, November 2026)
# Build (see Code/Install/Install-Library.md):
#   cmake -S . -B build
#   cmake --build build
# The programs are then in build/Examples/<example>/, with the name of the example. Requires Qt 6 (Widgets, Network, SerialPort) and C++17.
# The library and every example can also be built separately, with the CMakeLists.txt in its own directory.

cmake_minimum_required(VERSION 3.16)
project(XpressNetLinux VERSION 2.0 LANGUAGES CXX)

add_subdirectory(Code/XpressNetQt)
add_subdirectory("Examples/XpressNet Basic Test")
add_subdirectory("Examples/XpressNet Test")
add_subdirectory("Examples/Z21 Basic Test")
add_subdirectory("Examples/XpressNet Multiple Instances")
add_subdirectory("Examples/RS-Bus Monitor")
