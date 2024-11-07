#!/bin/bash
# Use the PORT environment variable if provided, or default to port 8501
PORT=${PORT:-8501}
streamlit run app.py --server.port $PORT --server.enableCORS false
