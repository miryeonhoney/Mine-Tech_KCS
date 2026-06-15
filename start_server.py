import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import dashboard_app
print("Server starting on http://127.0.0.1:8080")
dashboard_app.app.run(host="0.0.0.0", port=8080, debug=False, use_reloader=False)
