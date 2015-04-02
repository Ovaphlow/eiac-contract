# -*- coding=UTF-8 -*-

from tornado.wsgi import WSGIContainer
from tornado.httpserver import HTTPServer
from tornado.ioloop import IOLoop

from contract import app

def run():
    http_server = HTTPServer(WSGIContainer(app))
    http_server.listen(app.config['port'])
    IOLoop.instance().start()


if __name__ == '__main__':
    run()
