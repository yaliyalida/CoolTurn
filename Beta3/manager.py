#!/usr/bin/python
# coding: utf-8
from flask import Flask, render_template, request
from application.excel import EXCEL

from flask_bootstrap import Bootstrap
from flask_cors import CORS

print({'a','b','c'})

app = Flask(__name__)
app.register_blueprint(EXCEL)


CORS(app, supports_credentials=True)
Bootstrap(app)


@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template('main.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

