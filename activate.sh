#!/bin/bash

# Activate Hermit environment
source .hermit/python/bin/activate

# Run the K4 script with provided arguments
python k4.py "$@" 