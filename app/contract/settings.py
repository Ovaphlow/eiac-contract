# -*- coding=UTF-8 -*-


class Config(object):
    DEBUG = False
    SECRET_KEY = '%@*&#()_IKHTYHJFKnhiwkre99idu8ueU'

    PORT = 8090

    DB_HOST = '127.0.0.1'
    DB_USER = 'cmtech'
    DB_PASSWORD = 'cmtech.1123'
    DB_NAME = 'emsdatabase'

    TEMPLATE_FILE_PATH = 'static\\contract'
    TEMPLATE_FILE_NAME = 'template.xls'
    SOFFICE_PATH = 'C:\\"Program Files (x86)"\\"LibreOffice 4"\\program\\soffice.exe'


class DevelConfig(Config):
    DEBUG = True

    DB_HOST = '125.211.221.215'
