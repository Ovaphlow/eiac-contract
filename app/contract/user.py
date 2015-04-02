# coding:utf-8

import gl

from flask import redirect, render_template, session, request
from sqlalchemy import text

from contract import app


@app.route('/')
def home():
    if 'user_id' not in session:
        return redirect('/login')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        print('login')
        sql = '''
select
    *
from
    `tbs001_user`
where
    LoginName = :acc
    and
    Password = :pwd
    and
    RoleId != :role '''
        param = {
            'acc': request.form['account'],
            'pwd': request.form['password'],
            'role': u'环评单位',
        }
        res = gl.db_engine.execute(text(sql), param)
        row = res.fetchone()
        res.close()
        print(row)
        if not row:
            return redirect('/login')
        session['user_id'] = row.ID
        session['user_name'] = row.UserName
        session['user_account'] = row.userid
        if row.RoleId == u'建设单位':
            return redirect('/login')
        else:
            return redirect('/')

    return render_template('login.html')
