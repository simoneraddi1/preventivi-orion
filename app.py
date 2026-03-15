from flask import Flask, request, jsonify, send_file
from docx import Document
import json, io, os, re

app = Flask(__name__)

@app.route('/preventivo', methods=['POST'])
def genera_preventivo():
    raw = request.get_data(as_text=True)
    
    # Rimuovi backtick e "json" se presenti
    clean = re.sub(r'```json|```', '', raw).strip()
    
    try:
        dati = json.loads(clean)
    except:
        return jsonify({'error': 'JSON non valido', 'raw': raw}), 400
    
    doc = Document()
    doc.add_heading('PREVENTIVO', 0)
    doc.add_paragraph(f"Indirizzo: {dati.get('indirizzo', '')}")
    doc.add_paragraph(f"Data: {dati.get('data', '')}")
    doc.add_paragraph('')
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = 'Descrizione'
    hdr[1].text = 'Persone'
    hdr[2].text = 'Ore'
    hdr[3].text = 'Prezzo/h'
    hdr[4].text = 'Totale'
    
    for riga in dati.get('righe', []):
        row = table.add_row().cells
        row[0].text = f"{riga.get('descrizione','')}\n{riga.get('dettaglio','')}"
        row[1].text = str(riga.get('persone', 1))
        row[2].text = str(riga.get('ore', 0))
        row[3].text = f"{riga.get('prezzo_ora', 0)} €"
        row[4].text = f"{riga.get('totale', 0)} €"
    
    r = table.add_row().cells
    r[3].text = 'Subtotale'
    r[4].text = f"{dati.get('subtotale', 0)} €"
    
    r = table.add_row().cells
    r[2].text = 'IVA'
    r[3].text = '22%'
    r[4].text = f"{dati.get('iva', 0)} €"
    
    r = table.add_row().cells
    r[3].text = 'Totale'
    r[4].text = f"{dati.get('totale', 0)} €"
    
    doc.add_paragraph('')
    doc.add_paragraph('Termini di pagamento - 30GG D.F.')
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    
    return send_file(buf, as_attachment=True, download_name='preventivo.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/')
def home():
    return 'Server preventivi attivo!'

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
