# -*- coding=UTF-8 -*-
from flask import Flask

import gl
import view

app = Flask(__name__)
app.host = '0.0.0.0'
app.debug = True
app.secret_key = 'Ovaphlow'

app.add_url_rule('/', view_func=view.Home.as_view('home'))
app.add_url_rule('/login', view_func=view.Login.as_view('login'))
app.add_url_rule('/logout', view_func=view.Logout.as_view('logout'))

app.add_url_rule('/generate_contract',
    view_func=view.GenerateContract.as_view('generate_contract'))
app.add_url_rule('/generate_pdf',
    view_func=view.GeneratePDF.as_view('generate_pdf'))

app.add_url_rule('/eia', view_func=view.EIA.as_view('eia'))

if __name__ == '__main__':
    app.run(port=8089)
