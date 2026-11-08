#!/bin/bash

# Jira Scanner Main Script
# =====================================================

echo "=============================================="
echo "🔍 JIRA/GIRA SCANNER"
echo "CPU- Task Scanner with Excel Export"
echo "=============================================="
echo ""

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Check if main.py exists
if [[ ! -f "main.py" ]]; then
    echo "❌ Error: main.py not found!"
    echo ""
    echo "Please install dependencies first:"
    echo "   pip install -r requirements.txt"
    exit 1
fi

# Check if config.yaml exists
if [[ ! -f "config.yaml" ]]; then
    echo "⚠️  Warning: config.yaml not found!"
    echo ""
    echo "Creating default config.yaml..."
    cat > config.yaml << 'CONFIGEOF'
# Jira/Gira Scanner Configuration
server:
  url: "https://your-jira-server.com"
  protocol: "https"
  
api:
  token: ""
  
project:
  key: "CPU"
  
scripts:
  base_path: "/home/user1/work/jirascanner/scripts/"
  extensions: [".py", ".sh", ".sql"]
  
excel:
  output_path: "/home/user1/work/jirascanner/output/"
  output_filename: "jira_report_{YYYY}_{MM}_{DD}.xlsx"

logging:
  level: "INFO"
CONFIGEOF
    echo "✅ Created default config.yaml"
    echo ""
fi

# Run the scanner
echo "Running Jira Scanner..."
echo "Looking for tasks with CPU- prefix..."
echo ""
python3 main.py

# Check exit code
if [[ $? -eq 0 ]]; then
    echo ""
    echo "✅ Report generation completed successfully!"
    echo ""
    echo "Output files:"
    ls -lh output/*.xlsx 2>/dev/null || echo "  (No Excel files found)"
else
    echo ""
    echo "❌ Report generation failed!"
    echo "Check error messages above."
fi

echo "=============================================="
