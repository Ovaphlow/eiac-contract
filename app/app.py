# -*- coding=UTF-8 -*-

from contract import app

app.run(debug=app.config['DEBUG'], port=app.config['PORT'])
