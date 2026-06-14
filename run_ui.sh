#!/usr/bin/env bash
# One-command launcher for the Streamlit UI - no manual venv activation needed.
#   bash ./run_ui.sh
set -e
root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$root"

py="$root/.venv/bin/python"
if [ ! -x "$py" ]; then
    echo "Virtual env not found at .venv/. Run ./setup.sh first." >&2
    exit 1
fi

if ! "$py" -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('streamlit') else 1)"; then
    echo "Streamlit not installed - installing it into the venv..."
    "$py" -m pip install streamlit
fi

echo "Starting the UI at http://localhost:8501  (press Ctrl+C to stop)"
exec "$py" -m streamlit run "$root/ui/app.py"
