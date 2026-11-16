#!/bin/bash
# Dừng OCR SORT BOOK

echo "========================================"
echo "  Stopping OCR Pipeline Services"
echo "========================================"

# Dừng API
echo "--> Sending TERM signal to FastAPI (Uvicorn) server..."
pkill -f "uvicorn app.main:app"
if [ $? -eq 0 ]; then
    echo "    Signal sent."
else
    echo "    FastAPI server not found."
fi

# Dừng Spellfix
echo "--> Sending TERM signal to Spellfix service..."
pkill -f "uvicorn app.spellfix_api:app"
if [ $? -eq 0 ]; then
    echo "    Signal sent."
else
    echo "    Spellfix service not found."
fi

echo ""
echo "----------------------------------------"
echo "All application services have been sent the stop signal."
echo "----------------------------------------"
