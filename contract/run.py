# -*- coding=UTF-8 -*-
from tornado.wsgi import WSGIContainer
from tornado.httpserver import HTTPServer
from tornado.ioloop import IOLoop
from app import app


def run():
    http_server = HTTPServer(WSGIContainer(app))
    http_server.listen(8089)
    IOLoop.instance().start()


if __name__ == '__main__':
    run()
