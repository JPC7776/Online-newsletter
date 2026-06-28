#!/bin/bash
echo "✦ Launching LodgeGate & RezXS Newsletter Studio..."
cd "$(dirname "$0")"
sleep 2 && open "http://127.0.0.1:5050" &
python3 app.py
