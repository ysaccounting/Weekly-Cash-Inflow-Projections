import io
import json
import os
import base64
import anthropic
from flask import Flask, request, jsonify, send_file, render_template

from reconcile import (
    match_files_to_payments,
    build_part1_df,
    build_part2_data,
    write_combined_xlsx,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB


@app.route('/')
def index():
    return render_template('index.html')


# -- Screenshot extraction ----------------------------------------------------
@app.route('/extract_payments', methods=['POST'])
def extract_payments():
    screenshot = request.files.get('screenshot')
    if not screenshot:
        return jsonify({'error': 'No screenshot uploaded'}), 400

    img_bytes = screenshot.read()
    img_b64   = base64.standard_b64encode(img_bytes).decode('utf-8')
    mime_type = screenshot.content_type or 'image/png'

    client = anthropic.Anthropic()
    message = client.messages.create(
        model='claude-opus-4-5',
        max_tokens=2048,
        messages=[{
            'role': 'user',
            'content': [
                {
                    'type': 'image',
                    'source': {'type': 'base64', 'media_type': mime_type, 'data': img_b64},
                },
                {
                    'type': 'text',
                    'text': (
                        'This is a screenshot of the StubHub payment portal. '
                        'Extract every row from the payments table. '
                        'Return ONLY a JSON array, no other text, no markdown fences. '
                        'Each element must have these exact keys: '
                        'payment_id (string), date (string in DD/MM/YYYY format), '
                        'proceeds (number), charges (number), credit (number). '
                        'Preserve negative signs on charges/credit where present. '
                        'Example: [{"payment_id":"65780380","date":"13/05/2026",'
                        '"proceeds":86439.66,"charges":0,"credit":0}]'
                    ),
                },
            ],
        }],
    )

    raw = message.content[0].text.strip()
    if raw.startswith('```'):
        raw = raw.split('\n', 1)[1] if '\n' in raw else raw[3:]
    raw = raw.rstrip('`').strip()

    try:
        rows = json.loads(raw)
    except Exception:
        return jsonify({'error': 'Could not parse rows from screenshot', 'raw': raw}), 422

    return jsonify({'rows': rows})


# -- Part 1: preview file-to-payment matches ----------------------------------
@app.route('/match', methods=['POST'])
def match():
    csv_files = {f.filename: f.read() for f in request.files.getlist('csv_files')}
    try:
        payment_rows = json.loads(request.form.get('payments', '[]'))
    except Exception:
        return jsonify({'error': 'Invalid payments JSON'}), 400

    matched, unmatched_files, unmatched_payments = match_files_to_payments(
        csv_files, payment_rows
    )
    return jsonify({
        'matched': {pid: {'file': fname, 'date': date}
                    for pid, (fname, date) in matched.items()},
        'unmatched_files':    unmatched_files,
        'unmatched_payments': unmatched_payments,
    })


# -- Full process: generate combined output xlsx ------------------------------
@app.route('/process', methods=['POST'])
def process():
    csv_files = {f.filename: f.read() for f in request.files.getlist('csv_files')}

    stubhub_invoice = request.files.get('stubhub_invoice')
    if not stubhub_invoice:
        return jsonify({'error': 'StubHub invoice details file is required'}), 400
    stubhub_invoice_bytes = stubhub_invoice.read()

    try:
        payment_rows = json.loads(request.form.get('payments', '[]'))
    except Exception:
        return jsonify({'error': 'Invalid payments JSON'}), 400

    inv_det_file = request.files.get('inv_det_file')
    inv_file     = request.files.get('inv_file')
    if not inv_det_file or not inv_file:
        return jsonify({'error': 'Both Network Invoice Detail and Invoices files are required'}), 400
    inv_det_bytes = inv_det_file.read()
    inv_bytes     = inv_file.read()

    matched, unmatched_files, unmatched_payments = match_files_to_payments(
        csv_files, payment_rows
    )
    if not matched:
        return jsonify({'error': 'No CSV files could be matched to payment IDs'}), 400

    part1_df   = build_part1_df(csv_files, matched, stubhub_invoice_bytes)
    part2_data = build_part2_data(inv_det_bytes, inv_bytes)
    xlsx_bytes, filename = write_combined_xlsx(part1_df, part2_data)

    from flask import make_response
    import urllib.parse
    response = make_response(xlsx_bytes)
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    encoded = urllib.parse.quote(filename)
    response.headers['Content-Disposition'] = f'attachment; filename="{filename}"; filename*=UTF-8\'\'{encoded}'
    return response


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
