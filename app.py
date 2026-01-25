"""
Translation Processor - Smoke Test
A minimal web server to test network accessibility
"""
from flask import Flask, render_template_string
import socket

app = Flask(__name__)

# Simple HTML template
HOME_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Translation Processor - Smoke Test</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 50px auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #2c3e50;
        }
        .success {
            color: #27ae60;
            font-size: 24px;
            margin: 20px 0;
        }
        .info {
            background-color: #e8f4f8;
            padding: 15px;
            border-radius: 4px;
            margin: 20px 0;
        }
        .info-item {
            margin: 10px 0;
        }
        .label {
            font-weight: bold;
            color: #34495e;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🎉 Translation Processor</h1>
        <div class="success">✓ Server is running successfully!</div>

        <div class="info">
            <div class="info-item">
                <span class="label">Server Host:</span> {{ hostname }}
            </div>
            <div class="info-item">
                <span class="label">Your IP:</span> {{ ip_address }}
            </div>
            <div class="info-item">
                <span class="label">Status:</span> Ready for testing
            </div>
        </div>

        <h2>Next Steps:</h2>
        <ol>
            <li>If you're seeing this page, the server is working on THIS machine ✓</li>
            <li>Now try accessing from another computer on your network using:
                <ul>
                    <li><code>http://{{ ip_address }}:5000</code></li>
                    <li>Or: <code>http://{{ hostname }}:5000</code></li>
                </ul>
            </li>
            <li>If other machines can see this page, we're ready to proceed! ✓</li>
            <li>If blocked, we'll need to check firewall/security settings or pivot to desktop app</li>
        </ol>
    </div>
</body>
</html>
"""

@app.route('/')
def home():
    """Main page showing server is running"""
    hostname = socket.gethostname()
    try:
        ip_address = socket.gethostbyname(hostname)
    except:
        ip_address = "Unable to detect IP"

    return render_template_string(
        HOME_TEMPLATE,
        hostname=hostname,
        ip_address=ip_address
    )

if __name__ == '__main__':
    # Run on all network interfaces so other machines can access
    print("\n" + "="*60)
    print("Translation Processor - Smoke Test Server")
    print("="*60)
    hostname = socket.gethostname()
    try:
        ip_address = socket.gethostbyname(hostname)
        print(f"\nServer starting on: {hostname} ({ip_address})")
        print(f"\nAccess from this machine:")
        print(f"  → http://localhost:5000")
        print(f"\nAccess from other machines on your network:")
        print(f"  → http://{ip_address}:5000")
        print(f"  → http://{hostname}:5000")
    except:
        print("\nUnable to detect IP address")
        print("Try: http://localhost:5000 on this machine")

    print("\nPress CTRL+C to stop the server")
    print("="*60 + "\n")

    app.run(host='0.0.0.0', port=5000, debug=True)
