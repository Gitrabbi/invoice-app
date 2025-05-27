#!/bin/bash
# Set default port if PORT isn't set or is empty
PORT=${PORT:-8501}

# Validate port is a number
if ! [[ "$PORT" =~ ^[0-9]+$ ]]; then
  echo "Warning: Invalid PORT '$PORT'. Using default 8501."
  PORT=8501
fi

# Debug output
echo "Starting Streamlit on port: $PORT"

# Run Streamlit
exec streamlit run invoice_app1.py --server.port=$PORT --server.address=0.0.0.0
