#!/bin/bash
# Build script for Cellify WASM module

set -e

cd "$(dirname "$0")"

echo "Building WASM module..."

if ! command -v wasm-pack &> /dev/null; then
    echo "wasm-pack not found. Installing..."
    cargo install wasm-pack
fi

wasm-pack build --target web --out-dir ../src/formats/xlsx/wasm --release

rm -f ../src/formats/xlsx/wasm/.gitignore
rm -f ../src/formats/xlsx/wasm/package.json
rm -f ../src/formats/xlsx/wasm/README.md

echo "WASM module built successfully!"
echo "Output: src/formats/xlsx/wasm/"
