#!/bin/bash
echo ""
echo "================================================"
echo "  Teams Chat Viewer - PST Conversion Server"
echo "================================================"
echo ""

# Check Python
if ! command -v python3 &>/dev/null; then
    echo " [ERROR] Python 3 not found."
    echo " Install it from https://python.org or run: brew install python"
    exit 1
fi

echo " Checking dependencies..."

# Try to install libpff-python
pip3 install libpff-python -q 2>/dev/null || pip install libpff-python -q 2>/dev/null

# Also try readpst via brew if on mac
if command -v brew &>/dev/null && ! command -v readpst &>/dev/null; then
    echo " Installing readpst via Homebrew..."
    brew install libpst -q
fi

# Ubuntu/Debian
if command -v apt-get &>/dev/null && ! command -v readpst &>/dev/null; then
    echo " Installing readpst..."
    sudo apt-get install -y pst-utils -q
fi

echo " Starting server..."
echo ""
python3 server.py
