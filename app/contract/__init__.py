# -*- coding=UTF-8 -*-

from flask import Flask

app = Flask(__name__)

app.config.from_object('contract.settings.DevelConfig')

from contract import user
