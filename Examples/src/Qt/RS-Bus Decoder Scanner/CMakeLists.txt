# XpressNetQt example RS-Bus Decoder Scanner (version 1.0, November 2026)
# Runs on Windows and Linux. Build: see Examples-Windows.md / Examples-Linux.md in the parent directory
# Requires Qt 6 (Widgets, Network, SerialPort) and a C++17 compiler.
# The CV overview is the same file as in the macOS version (../../macOS/RS-Bus Decoder Scanner/.../Helpfiles).

cmake_minimum_required(VERSION 3.16)
project(RsBusDecoderScanner VERSION 1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

find_package(Qt6 REQUIRED COMPONENTS Widgets)

# Visual Studio: the source files are UTF-8
if(MSVC)
  add_compile_options(/utf-8)
endif()

# The library. Not needed if this example is built from Examples/src/Qt/CMakeLists.txt (the library is then built there)
if(NOT TARGET XpressNetQt)
  add_subdirectory(../../../../Library/src/Qt/XpressNetQt XpressNetQt)
endif()

add_executable(RsBusDecoderScanner WIN32 MACOSX_BUNDLE
  main.cpp
  MainWindow.h
  MainWindow.cpp
  DecoderScanner.h
  DecoderScanner.cpp
  PreferencesDialog.h
  PreferencesDialog.cpp
  Settings.h
  Settings.cpp
  resources/resources.qrc
)

# Windows: the icon of the program
if(WIN32)
  target_sources(RsBusDecoderScanner PRIVATE resources/RS-Bus-Decoder-Scanner.rc)
endif()

# Visual Studio: while debugging the Qt DLLs are found via PATH
if(MSVC)
  set_target_properties(RsBusDecoderScanner PROPERTIES VS_DEBUGGER_ENVIRONMENT "PATH=$<TARGET_FILE_DIR:Qt6::Core>;%PATH%")
endif()

# The CV overview, as resource with prefix /help
set(HELP_DIR "${CMAKE_CURRENT_SOURCE_DIR}/../../macOS/RS-Bus Decoder Scanner/RS-Bus Decoder Scanner/Helpfiles")
qt_add_resources(RsBusDecoderScanner "help" PREFIX "/help" BASE "${HELP_DIR}" FILES "${HELP_DIR}/CV-Overview.html")

target_link_libraries(RsBusDecoderScanner PRIVATE Qt6::Widgets XpressNetQt)
target_compile_definitions(RsBusDecoderScanner PRIVATE APP_VERSION="${PROJECT_VERSION}")
# The name of the program, the same as in the macOS version
set_target_properties(RsBusDecoderScanner PROPERTIES OUTPUT_NAME "RS-Bus Decoder Scanner")
