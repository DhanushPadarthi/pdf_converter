#!/usr/bin/env bash
# Exit on error
set -o errexit

# Update package lists
apt-get update

# Install system dependencies for OCR
apt-get install -y 
    tesseract-ocr 
    tesseract-ocr-eng 
    tesseract-ocr-deu 
    tesseract-ocr-fra 
    tesseract-ocr-spa 
    poppler-utils 
    libtesseract-dev 
    libleptonica-dev 
    libgl1-mesa-glx 
    libglib2.0-0

# Upgrade pip and setuptools
pip install --upgrade pip setuptools wheel

# Install Python dependencies
pip install -r requirements.txt

echo "Build completed successfully!"
