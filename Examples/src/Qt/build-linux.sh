#!/bin/sh
# XpressNet Qt examples: build all examples for Linux (November 2026)
# Start in a terminal: ./build-linux.sh (or: sh build-linux.sh). See Examples-Linux.md.
# Builds the library and the six examples (Release), and makes for every example a zip file with the program
# in ../../compiled/Linux-Qt. The directory build is first removed.
# Requires Qt 6 (Widgets, Network, SerialPort), CMake, a C++17 compiler and zip.

set -e
cd "$(dirname "$0")"

# The library source must be found relative to this directory (the complete repository must be unzipped)
if [ ! -f ../../../Library/src/Qt/XpressNetQt/CMakeLists.txt ]; then
  echo "ERROR: ../../../Library/src/Qt/XpressNetQt not found."
  echo "Unzip the complete repository and start this file in its directory Examples/src/Qt."
  exit 1
fi
for tool in cmake zip; do
  if ! command -v $tool > /dev/null; then
    echo "ERROR: $tool not found (see Library/src/Qt/Install/Install-Library.md; zip: sudo apt install zip)."
    exit 1
  fi
done

# Build the library and all examples
rm -rf build
cmake -S . -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build --parallel

# One zip file per example
mkdir -p ../../compiled/Linux-Qt
for example in "XpressNet Basic Test" "XpressNet Test" "Z21 Basic Test" "XpressNet Multiple Instances" \
               "RS-Bus Monitor" "RS-Bus Decoder Scanner"; do
  echo "==== $example"
  rm -f "../../compiled/Linux-Qt/$example.zip"
  (cd "build/$example" && zip -q "../../../../compiled/Linux-Qt/$example.zip" "$example")
done

echo
echo "Ready. The zip files are in Examples/compiled/Linux-Qt, the programs in build/<example>/."
