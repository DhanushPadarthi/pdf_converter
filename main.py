import os

# This tells Replit it's a Python Flask web app
if __name__ == "__main__":
    from app import app
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
