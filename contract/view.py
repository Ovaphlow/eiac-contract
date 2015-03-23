# -*- coding=UTF-8 -*-

import os
import sys

from datetime import datetime

from flask import redirect, render_template, session, request
from flask.views import MethodView
from sqlalchemy import text

from gl import *

reload(sys)
sys.setdefaultencoding("utf-8")


class GeneratePDF(MethodView):
    def get(self):
        if not 'user_id' in session:
            return redirect('/login')
        project_id = request.args.get('project_id')
        generate_pdf(project_id)
        return redirect('/static/contract/%s.pdf' % (project_id))


class EIA(MethodView):
    def get(self):
        if not 'user_id' in session:
            return redirect('/login')
        name = request.args.get('name', '')
        sql = '''
            select *
            from tbs001_developprojectbasicinfo
            where projectreportstate=1
            and noticestate!=2
            and (shoulituihuishanchu=1
                or shoulituihuishanchu is null)
            and locate(:name,ProjectName)>0
            order by id desc
            limit 20
        '''
        param = {'name': name}
        res = db_engine.execute(text(sql), param)
        rows = res.fetchall()
        res.close()
        return render_template('eia.html',
            user_name=session['user_name'], rows=rows)

    def post(self):
        name = request.form['name']
        return redirect('/eia?name=%s' % (name))


class GenerateContract(MethodView):
    def get(self):
        if not 'user_id' in session:
            return redirect('/login')
        project_id = request.args.get('project_id')
        generate_contract_1(project_id)
        return redirect('/static/contract/%s.xls' % (project_id))


class Home(MethodView):
    def get(post):
        if not 'user_id' in session:
            return redirect('/login')
        sql = '''
            select *
            from tbs001_developprojectbasicinfo
            where js=:user_id
            and projectreportstate=1
            and noticestate!=2
            and (shoulituihuishanchu=1
                or shoulituihuishanchu is null)
            order by id desc
        '''
        param = {'user_id': session['user_account']}
        res = db_engine.execute(text(sql), param)
        rows = res.fetchall()
        res.close()
        return render_template('home.html',
            user_name=session['user_name'], rows=rows)


class Login(MethodView):
    def get(self):
        return render_template('login.html')

    def post(self):
        """
        2015-03-16
        常晓宇
        禁止环评单位及建设单位下载合同
        if row.RoleId == u'建设单位':
            return redirect('/login')
        原为
        if row.RoleId == u'建设单位':
            return redirect('/')
        """
        sql = '''
select *
from `tbs001_user`
where LoginName=:acc
and Password=:pwd
and RoleId!=:role
        '''
        param = {
            'acc': request.form['account'],
            'pwd': request.form['password'],
            'role': u'环评单位',
        }
        res = db_engine.execute(text(sql), param)
        row = res.fetchone()
        res.close()
        if not row:
            return redirect('/login')
        session['user_id'] = row.ID
        session['user_name'] = row.UserName
        session['user_account'] = row.userid
        if row.RoleId == u'建设单位':
            return redirect('/login')
        else:
            return redirect('/eia')


class Logout(MethodView):
    def get(self):
        session.pop('user_id', None)
        session.pop('user_name', None)
        session.pop('user_account', None)
        return redirect('/login')
