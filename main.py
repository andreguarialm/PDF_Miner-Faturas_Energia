import os
import urllib
from flask import Flask, make_response, jsonify, request
from flask_cors import CORS
from PDFMinerFINAL import capture_invoice

app = Flask('eltanin-miner')
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1000 * 1000 # 10MB

@app.route('/', methods=['GET'])
def ping():
    response = jsonify({ 'server': 'Flask app', 'name': 'eltanin-miner' })
    return make_response(response)

@app.route('/capture', methods=['POST'])
def capture():
    distributor = request.json['distributorId']
    file = request.json['file']
    if (file):
        resp = urllib.request.urlopen(file)
        with open('invoice.pdf', 'wb') as f:
            f.write(resp.file.read())
    else:
        open('invoice.pdf', 'wb')

    result_capture = capture_invoice()

    response = jsonify({'data': result_capture})
    return make_response(response)

app.run(debug=True)