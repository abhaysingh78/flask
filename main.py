from flask import Flask, stream_with_context, Response
from flask_cors import CORS
import pandas as pd
import json 
import time

app = Flask(__name__)
CORS(app)
@app.route('/')
def generate_json_data():
        while True:
            json = pd.read_csv('./optchain.csv')
            json_data = json.to_json(orient='records')
            # print(json_data)
            # yield json_data
            # time.sleep(2)
            return json_data
            
        
@app.route('/data', methods=['GET'])
def stream():
    return Response(stream_with_context(generate_json_data()), content_type='application/json')

if __name__ == '__main__':
    app.run(debug=True, port=8000)