from flask import Flask, request, jsonify, send_file
from docx import Document
import json, io, os, re

app = Flask(__name__)

@app.route('/preventivo', methods=['POST'])
def genera_preventivo():
    raw = request.get_data(as_text=True)
    clean = re.sub(r'```json|```', '', raw).strip()

    try:
        dati = json.loads(clean)
    except:
        return jsonify({'error': 'JSON non valido', 'raw': raw}), 400

    return jsonify({"ok": True, "dati": dati}), 200

@app.route('/')
def home():
    return 'Server preventivi attivo!'

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
