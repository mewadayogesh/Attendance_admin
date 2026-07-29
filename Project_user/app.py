import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import io
import os
import traceback
import calendar as cal
from datetime import datetime
from functools import wraps

# Loads variables from a local .env file (if present) into os.environ, so
# MAIL_SENDER_EMAIL / MAIL_SENDER_PASSWORD / MAIL_RECIPIENT_EMAIL etc. can be
# set once in .env instead of exporting them in every terminal session.
# In production (Render/Heroku/etc.) this is a harmless no-op — set the real
# variables in that platform's dashboard instead of relying on a .env file.
# Requires: pip install python-dotenv
from dotenv import load_dotenv
load_dotenv()

#this for templates or html page adding by flask 
from flask import (
    Flask, render_template, request, redirect, url_for, session,
    flash, send_file, abort
)
from openpyxl import Workbook
import db

app = Flask(__name__, static_folder='static', instance_relative_config=True)

# NOTE: pull the secret key from the environment in production. Falling back
# to a hardcoded string is fine for local dev only.
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'crest@#2026')

os.makedirs(app.instance_path, exist_ok=True)
app.config['DATABASE'] = os.path.join(app.instance_path, 'data.db')

db.init_app(app)

# ---------------------------------------------------------------------------
# Email configuration
# ---------------------------------------------------------------------------
MAIL_SENDER_EMAIL = os.environ.get('MAIL_SENDER_EMAIL', '')
MAIL_SENDER_PASSWORD = os.environ.get('MAIL_SENDER_PASSWORD', '')
MAIL_RECIPIENT_EMAIL = os.environ.get('MAIL_RECIPIENT_EMAIL', '')


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return view(*args, **kwargs)
    return wrapped_view


def roles_required(*allowed_roles):
    def decorator(view):
        @wraps(view)
        def wrapped_view(*args, **kwargs):
            if not session.get('logged_in'):
                return redirect(url_for('login'))
            if session.get('role') not in allowed_roles:
                abort(403)
            return view(*args, **kwargs)
        return wrapped_view
    return decorator


def build_xlsx_buffer(headers, rows):
    """Build an in-memory .xlsx workbook and return its BytesIO buffer."""
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)

    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)
    for col_cells in ws.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def make_xlsx_response(headers, rows, filename):
    buffer = build_xlsx_buffer(headers, rows)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


def format_date_ddmmyyyy(value):
    if not value:
        return value
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d-%m-%Y')
    except (ValueError, TypeError):
        return value


def format_datetime_ddmmyyyy(value):
    if not value:
        return value
    try:
        return datetime.strptime(value, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')
    except (ValueError, TypeError):
        return value


def get_employee_or_404(employee_id):
    conn = db.get_db()
    row = conn.execute('SELECT * FROM employees WHERE id = ?', (employee_id,)).fetchone()
    return row


def get_leave_or_404(request_id):
    conn = db.get_db()
    row = conn.execute('SELECT * FROM leave_requests WHERE id = ?', (request_id,)).fetchone()
    return row


def get_attendance_or_404(attendance_id):
    conn = db.get_db()
    row = conn.execute('SELECT * FROM attendance WHERE id = ?', (attendance_id,)).fetchone()
    return row


def get_current_employee():
    username = session.get('username')
    if not username:
        return None
    conn = db.get_db()
    user_row = conn.execute('SELECT linked_employee_id FROM users WHERE username = ?', (username,)).fetchone()
    if user_row and user_row['linked_employee_id']:
        emp = conn.execute('SELECT * FROM employees WHERE id = ?', (user_row['linked_employee_id'],)).fetchone()
        if emp:
            return emp
    return conn.execute(
        'SELECT * FROM employees WHERE name = ? OR employee_id = ?', (username, username)
    ).fetchone()


ATTENDANCE_STATUSES = ['Present', 'Half Day', 'Leave', 'Absent']


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

@app.route('/', methods=['GET'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('dashboard'))
    return render_template('login.html', error=None)


@app.route('/login', methods=['POST'])
def do_login():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()

    if not username or not password:
        return render_template('login.html', error='Please fill out this field.')

    conn = db.get_db()
    user_record = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()

    if user_record and user_record['password'] == password:
        session['logged_in'] = True
        session['username'] = username
        session['role'] = user_record['role']
        return redirect(url_for('dashboard'))

    return render_template('login.html', error='Invalid Username or Password')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

@app.route('/dashboard')
@login_required
def dashboard():
    conn = db.get_db()

    employee_count = conn.execute('SELECT COUNT(*) AS c FROM employees').fetchone()['c']

    if session.get('role') == 'user':
        leave_count = conn.execute('SELECT COUNT(*) AS c FROM leave_requests WHERE submitted_by = ?', (session.get('username'),)).fetchone()['c']
        pending_leave_count = conn.execute("SELECT COUNT(*) AS c FROM leave_requests WHERE submitted_by = ? AND status = 'Pending'", (session.get('username'),)).fetchone()['c']
        recent_leaves = conn.execute('SELECT * FROM leave_requests WHERE submitted_by = ? ORDER BY id DESC LIMIT 5', (session.get('username'),)).fetchall()
    else:
        leave_count = conn.execute('SELECT COUNT(*) AS c FROM leave_requests').fetchone()['c']
        pending_leave_count = conn.execute("SELECT COUNT(*) AS c FROM leave_requests WHERE status = 'Pending'").fetchone()['c']
        recent_leaves = conn.execute('SELECT * FROM leave_requests ORDER BY id DESC LIMIT 5').fetchall()

    recent_employees = conn.execute('SELECT * FROM employees ORDER BY id DESC LIMIT 5').fetchall()

    my_employee = get_current_employee()
    my_attendance_today = None
    if my_employee:
        today = datetime.now().strftime('%Y-%m-%d')
        row = conn.execute(
            'SELECT * FROM attendance WHERE employee_id = ? AND work_date = ?', (my_employee['id'], today)
        ).fetchone()
        my_attendance_today = row['status'] if row else None

    return render_template(
        'dashboard.html',
        session_username=session.get('username'),
        session_role=session.get('role'),
        employee_count=employee_count,
        leave_count=leave_count,
        pending_leave_count=pending_leave_count,
        recent_employees=recent_employees,
        recent_leaves=recent_leaves,
        my_employee=my_employee,
        my_attendance_today=my_attendance_today
    )


# ---------------------------------------------------------------------------
# user management
# ---------------------------------------------------------------------------

@app.route('/add_user', methods=['GET', 'POST'])
@roles_required('admin')
def add_user():
    conn = db.get_db()
    employees = conn.execute('SELECT * FROM employees ORDER BY name ASC').fetchall()

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        role = request.form.get('role', 'user').strip()
        linked_employee_id = request.form.get('linked_employee_id', '').strip() or None

        if not username or not password:
            flash('Username and password are required.', 'error')
            return render_template('add_user.html', employees=employees)

        existing_user = conn.execute('SELECT id FROM users WHERE username = ?', (username,)).fetchone()
        if existing_user:
            flash('Username already exists. Choose a different one.', 'error')
            return render_template('add_user.html', employees=employees)

        local_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn.execute(
            'INSERT INTO users (username, password, role, linked_employee_id, created_at) VALUES (?, ?, ?, ?, ?)',
            (username, password, role, linked_employee_id, local_timestamp)
        )
        conn.commit()
        flash('New user created successfully!', 'success')
        return redirect(url_for('dashboard'))

    return render_template('add_user.html', employees=employees)


@app.route('/report/users')
@roles_required('admin')
def report_users():
    conn = db.get_db()
    users = conn.execute('SELECT * FROM users ORDER BY id ASC').fetchall()
    return render_template('report_users.html', users=users, user_role=session.get('role'))


@app.route('/users/<int:user_id>/edit', methods=['GET', 'POST'])
@roles_required('admin')
def edit_user(user_id):
    conn = db.get_db()
    user = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
    employees = conn.execute('SELECT * FROM employees ORDER BY name ASC').fetchall()

    if user is None:
        flash('That user no longer exists.', 'error')
        return redirect(url_for('report_users'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        role = request.form.get('role', 'user').strip()
        linked_employee_id = request.form.get('linked_employee_id', '').strip() or None

        if not username or not password:
            flash('Username and password are required.', 'error')
            return render_template('user_edit.html', user=user, employees=employees)

        conn.execute(
            'UPDATE users SET username = ?, password = ?, role = ?, linked_employee_id = ? WHERE id = ?',
            (username, password, role, linked_employee_id, user_id)
        )
        conn.commit()
        flash('User record updated successfully.', 'success')
        return redirect(url_for('report_users'))

    return render_template('user_edit.html', user=user, employees=employees)


@app.route('/users/<int:user_id>/delete', methods=['POST'])
@roles_required('admin')
def delete_user(user_id):
    conn = db.get_db()
    conn.execute('DELETE FROM users WHERE id = ?', (user_id,))

    remaining = conn.execute('SELECT id FROM users ORDER BY id ASC').fetchall()
    conn.execute('DELETE FROM sqlite_sequence WHERE name="users"')
    for new_id, row in enumerate(remaining, start=1):
        conn.execute('UPDATE users SET id = ? WHERE id = ?', (new_id, row['id']))
        conn.execute('INSERT INTO sqlite_sequence (name, seq) VALUES ("users", ?)', (new_id,))

    conn.commit()
    flash('User deleted successfully.', 'success')
    return redirect(url_for('report_users'))


# ---------------------------------------------------------------------------
# Employees
# ---------------------------------------------------------------------------

@app.route('/new_entry', methods=['GET', 'POST'])
@roles_required('admin', 'editor')
def new_entry():
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        employee_id = request.form.get('employee_id', '').strip()
        designation = request.form.get('designation', '').strip()
        dob = request.form.get('dob', '').strip()
        date_of_joining = request.form.get('date_of_joining', '').strip()

        if not name or not employee_id or not designation or not dob or not date_of_joining:
            return render_template(
                'new_entry.html',
                error_message="All input fields are required.",
                name=name, employee_id=employee_id, designation=designation, dob=dob, date_of_joining=date_of_joining
            )

        conn = db.get_db()
        existing_employee = conn.execute(
            'SELECT id FROM employees WHERE employee_id = ?', (employee_id,)
        ).fetchone()

        if existing_employee:
            return render_template(
                'new_entry.html',
                error_message="This Employee ID already exists. Please use a unique Employee ID.",
                name=name, employee_id=employee_id, designation=designation, dob=dob, date_of_joining=date_of_joining
            )

        local_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        conn.execute(
            'INSERT INTO employees (name, employee_id, designation, dob, date_of_joining, created_at) '
            'VALUES (?, ?, ?, ?, ?, ?)',
            (name, employee_id, designation, dob, date_of_joining, local_timestamp),
        )
        conn.commit()

        flash('Employee record saved successfully.', 'success')
        return redirect(url_for('report_employees'))

    return render_template('new_entry.html')


@app.route('/search', methods=['GET'])
@login_required
def search():
    query = request.args.get('employee_id', '').strip()
    results = []
    if query:
        conn = db.get_db()
        results = conn.execute(
            'SELECT * FROM employees WHERE employee_id LIKE ? ORDER BY id DESC',
            (f'%{query}%',),
        ).fetchall()
    return render_template('search.html', query=query, results=results, user_role=session.get('role'))


@app.route('/employees/<int:employee_id>/edit', methods=['GET', 'POST'])
@roles_required('admin', 'editor')
def edit_employee(employee_id):
    employee = get_employee_or_404(employee_id)
    if employee is None:
        flash('That employee record no longer exists.', 'error')
        return redirect(url_for('report_employees'))

    back_url = request.values.get('next') or url_for('report_employees')

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        emp_code = request.form.get('employee_id', '').strip()
        designation = request.form.get('designation', '').strip()
        dob = request.form.get('dob', '').strip()
        date_of_joining = request.form.get('date_of_joining', '').strip()

        if not name or not emp_code:
            flash('Name and Employee ID are required.', 'error')
            return redirect(url_for('edit_employee', employee_id=employee_id, next=back_url))

        conn = db.get_db()
        conn.execute(
            'UPDATE employees SET name = ?, employee_id = ?, designation = ?, dob = ?, '
            'date_of_joining = ? WHERE id = ?',
            (name, emp_code, designation, dob, date_of_joining, employee_id),
        )
        conn.commit()
        flash('Employee record updated.', 'success')
        return redirect(back_url)

    return render_template('employee_edit.html', employee=employee, back_url=back_url)


@app.route('/employees/<int:employee_id>/delete', methods=['POST'])
@roles_required('admin', 'editor')
def delete_employee(employee_id):
    conn = db.get_db()
    conn.execute('DELETE FROM employees WHERE id = ?', (employee_id,))

    remaining = conn.execute('SELECT id FROM employees ORDER BY id ASC').fetchall()
    conn.execute('DELETE FROM sqlite_sequence WHERE name="employees"')
    for new_id, row in enumerate(remaining, start=1):
        conn.execute('UPDATE employees SET id = ? WHERE id = ?', (new_id, row['id']))
        conn.execute('INSERT INTO sqlite_sequence (name, seq) VALUES ("employees", ?)', (new_id,))

    conn.commit()
    return redirect(request.form.get('next') or url_for('report_employees'))


# ---------------------------------------------------------------------------
# Leave requests
# ---------------------------------------------------------------------------

@app.route('/leave', methods=['GET', 'POST'])
@login_required
def leave():
    conn = db.get_db()
    username = session.get('username')
    emp_record = conn.execute('SELECT employee_id FROM employees WHERE name = ? OR employee_id = ?', (username, username)).fetchone()
    auto_employee_id = emp_record['employee_id'] if emp_record else ''

    if request.method == 'POST':
        dates = request.form.get('dates', '').strip()
        num_days = request.form.get('num_days', '').strip()
        reason = request.form.get('reason', '').strip()
        description = request.form.get('description', '').strip()
        employee_id = request.form.get('employee_id', '').strip() or auto_employee_id
        request_date = request.form.get('request_date', '').strip()

        if not dates or not num_days or not employee_id or not request_date:
            flash('All required fields must be filled out.', 'error')
            return render_template('leave.html', employee_id=auto_employee_id, dates=dates, num_days=num_days, reason=reason, request_date=request_date)

        local_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        conn.execute(
            'INSERT INTO leave_requests (employee_id, dates, num_days, reason, description, request_date, '
            'status, submitted_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
            (employee_id, dates, num_days, reason, description, request_date,
             'Pending', session.get('username'), local_timestamp),
        )
        conn.commit()
        flash('Leave request submitted successfully.', 'success')
        return redirect(url_for('report_leave'))

    return render_template('leave.html', employee_id=auto_employee_id)


@app.route('/leave/<int:request_id>/edit', methods=['GET', 'POST'])
@roles_required('admin', 'editor')
def edit_leave(request_id):
    leave_request = get_leave_or_404(request_id)
    if leave_request is None:
        flash('That leave request no longer exists.', 'error')
        return redirect(url_for('report_leave'))

    if request.method == 'POST':
        dates = request.form.get('dates', '').strip()
        num_days = request.form.get('num_days', '').strip()
        reason = request.form.get('reason', '').strip()
        description = request.form.get('description', '').strip()
        employee_id = request.form.get('employee_id', '').strip()
        request_date = request.form.get('request_date', '').strip()
        status = request.form.get('status', 'Pending').strip()

        if not dates or not num_days or not employee_id or not request_date:
            flash('All required fields must be filled out.', 'error')
            return redirect(url_for('edit_leave', request_id=request_id))

        conn = db.get_db()
        conn.execute(
            'UPDATE leave_requests SET employee_id = ?, dates = ?, num_days = ?, reason = ?, description = ?, request_date = ?, '
            'status = ? WHERE id = ?',
            (employee_id, dates, num_days, reason, description, request_date, status, request_id),
        )
        conn.commit()
        flash('Leave request updated.', 'success')
        return redirect(url_for('report_leave'))

    return render_template('leave_edit.html', leave_request=leave_request)


@app.route('/leave/<int:request_id>/delete', methods=['POST'])
@roles_required('admin', 'editor')
def delete_leave(request_id):
    conn = db.get_db()
    conn.execute('DELETE FROM leave_requests WHERE id = ?', (request_id,))

    remaining = conn.execute('SELECT id FROM leave_requests ORDER BY id ASC').fetchall()
    conn.execute('DELETE FROM sqlite_sequence WHERE name="leave_requests"')
    for new_id, row in enumerate(remaining, start=1):
        conn.execute('UPDATE leave_requests SET id = ? WHERE id = ?', (new_id, row['id']))
        conn.execute('INSERT INTO sqlite_sequence (name, seq) VALUES ("leave_requests", ?)', (new_id,))

    conn.commit()
    return redirect(url_for('report_leave'))


@app.route('/report/leave/<int:request_id>/<action>', methods=['POST'])
@roles_required('admin', 'editor')
def update_leave_status(request_id, action):
    if action not in ('approve', 'reject'):
        flash('Unknown action.', 'error')
        return redirect(url_for('report_leave'))

    conn = db.get_db()
    if action == 'approve':
        conn.execute('UPDATE leave_requests SET status = ? WHERE id = ?', ('Approved', request_id))
    else:
        description = request.form.get('description', '').strip()
        conn.execute('UPDATE leave_requests SET status = ?, description = ? WHERE id = ?', ('Rejected', description, request_id))

    conn.commit()
    flash(f'Leave request #{request_id} updated successfully.', 'success')
    return redirect(url_for('report_leave'))


# ---------------------------------------------------------------------------
# Attendance
# ---------------------------------------------------------------------------

@app.route('/attendance', methods=['GET'])
@login_required
def attendance_home():
    employee = get_current_employee()
    if employee is None:
        flash('Your account isn\'t linked to an employee record yet — ask an admin to link it from Add User / Edit User.', 'error')
        return redirect(url_for('dashboard'))

    conn = db.get_db()
    today = datetime.now().strftime('%Y-%m-%d')
    today_row = conn.execute(
        'SELECT * FROM attendance WHERE employee_id = ? AND work_date = ?', (employee['id'], today)
    ).fetchone()
    recent = conn.execute(
        'SELECT * FROM attendance WHERE employee_id = ? ORDER BY work_date DESC LIMIT 10', (employee['id'],)
    ).fetchall()

    return render_template('attendance.html', today_row=today_row, recent=recent, today=today, employee=employee)


@app.route('/attendance/check-in', methods=['POST'])
@login_required
def attendance_check_in():
    employee = get_current_employee()
    if employee is None:
        return {'error': 'no employee record linked to this account'}, 400

    conn = db.get_db()
    today = datetime.now().strftime('%Y-%m-%d')
    now = datetime.now().strftime('%H:%M:%S')

    existing = conn.execute(
        'SELECT * FROM attendance WHERE employee_id = ? AND work_date = ?', (employee['id'], today)
    ).fetchone()
    if existing:
        return {'error': 'already checked in today'}, 400

    conn.execute(
        'INSERT INTO attendance (employee_id, work_date, check_in, status) VALUES (?, ?, ?, ?)',
        (employee['id'], today, now, 'Present')
    )
    conn.commit()
    return {'ok': True, 'check_in': now}


@app.route('/attendance/check-out', methods=['POST'])
@login_required
def attendance_check_out():
    employee = get_current_employee()
    if employee is None:
        return {'error': 'no employee record linked to this account'}, 400

    conn = db.get_db()
    today = datetime.now().strftime('%Y-%m-%d')
    now = datetime.now().strftime('%H:%M:%S')

    existing = conn.execute(
        'SELECT * FROM attendance WHERE employee_id = ? AND work_date = ?', (employee['id'], today)
    ).fetchone()
    if not existing:
        return {'error': 'check in first'}, 400
    if existing['check_out']:
        return {'error': 'already checked out today'}, 400

    conn.execute('UPDATE attendance SET check_out = ? WHERE id = ?', (now, existing['id']))
    conn.commit()
    return {'ok': True, 'check_out': now}


@app.route('/attendance/calendar')
@login_required
def attendance_calendar():
    conn = db.get_db()
    now = datetime.now()
    year = int(request.args.get('year', now.year))
    month = int(request.args.get('month', now.month))

    employees = []
    viewing_id = None
    viewing_name = None

    if session.get('role') in ('admin', 'editor'):
        employees = conn.execute('SELECT * FROM employees ORDER BY name ASC').fetchall()
        requested = request.args.get('employee_id')
        if requested:
            viewing_id = int(requested)
        elif employees:
            viewing_id = employees[0]['id']
    else:
        my_employee = get_current_employee()
        if my_employee:
            viewing_id = my_employee['id']

    day_status = {}
    if viewing_id:
        rows = conn.execute(
            "SELECT work_date, status, check_in, check_out FROM attendance "
            "WHERE employee_id = ? AND strftime('%Y', work_date) = ? AND strftime('%m', work_date) = ?",
            (viewing_id, str(year), f'{month:02d}')
        ).fetchall()
        day_status = {r['work_date']: dict(r) for r in rows}
        emp_row = conn.execute('SELECT name FROM employees WHERE id = ?', (viewing_id,)).fetchone()
        viewing_name = emp_row['name'] if emp_row else None
    elif session.get('role') not in ('admin', 'editor'):
        flash('Your account isn\'t linked to an employee record yet — ask an admin to link it from Add User / Edit User.', 'error')

    holiday_rows = conn.execute(
        "SELECT holiday_date, name FROM holidays "
        "WHERE strftime('%Y', holiday_date) = ? AND strftime('%m', holiday_date) = ?",
        (str(year), f'{month:02d}')
    ).fetchall()
    holiday_map = {r['holiday_date']: r['name'] for r in holiday_rows}

    cal.setfirstweekday(cal.SUNDAY)
    weeks = []
    weekend_count = 0
    for week in cal.monthcalendar(year, month):
        week_cells = []
        for day_num in week:
            if day_num == 0:
                week_cells.append(None)
                continue
            iso = f'{year:04d}-{month:02d}-{day_num:02d}'
            info = day_status.get(iso)
            holiday_name = holiday_map.get(iso)
            is_sunday = datetime(year, month, day_num).weekday() == 6

            if info:
                status_label = info['status']
                css = (info['status'] or '').lower().replace(' ', '')
                extra_tag = None
                if holiday_name:
                    extra_tag = f'Worked on Holiday ({holiday_name})'
                elif is_sunday:
                    extra_tag = 'Worked on Sunday'
            elif holiday_name:
                status_label = 'Holiday'
                css = 'holiday'
                extra_tag = None
            elif is_sunday:
                status_label = 'Weekend'
                css = 'weekend'
                extra_tag = None
                weekend_count += 1
            else:
                status_label = None
                css = ''
                extra_tag = None

            week_cells.append({
                'day': day_num,
                'iso': iso,
                'status': status_label,
                'css': css,
                'holiday_name': holiday_name,
                'extra_tag': extra_tag,
                'is_sunday': is_sunday,
                'check_in': info['check_in'] if info else None,
                'check_out': info['check_out'] if info else None,
                'is_today': iso == now.strftime('%Y-%m-%d'),
            })
        weeks.append(week_cells)

    prev_month = month - 1 or 12
    prev_year = year - 1 if month == 1 else year
    next_month = month % 12 + 1
    next_year = year + 1 if month == 12 else year

    counts = {'Present': 0, 'Half Day': 0, 'Leave': 0, 'Absent': 0}
    for info in day_status.values():
        if info['status'] in counts:
            counts[info['status']] += 1
    holiday_count = len(holiday_map)

    return render_template(
        'calendar.html',
        weeks=weeks,
        month_name=cal.month_name[month],
        year=year, month=month,
        prev_month=prev_month, prev_year=prev_year,
        next_month=next_month, next_year=next_year,
        employees=employees, viewing_id=viewing_id, viewing_name=viewing_name,
        counts=counts, holiday_count=holiday_count, weekend_count=weekend_count,
        weekday_labels=['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    )


# ---------------------------------------------------------------------------
# Holidays
# ---------------------------------------------------------------------------

@app.route('/holidays', methods=['GET'])
@login_required
def holidays():
    conn = db.get_db()
    all_holidays = conn.execute('SELECT * FROM holidays ORDER BY holiday_date ASC').fetchall()
    return render_template('holidays.html', holidays=all_holidays, user_role=session.get('role'))


@app.route('/holidays/add', methods=['POST'])
@roles_required('admin')
def add_holiday():
    holiday_date = request.form.get('holiday_date', '').strip()
    name = request.form.get('name', '').strip()

    if not holiday_date or not name:
        flash('Date and name are required to add a holiday.', 'error')
        return redirect(url_for('holidays'))

    conn = db.get_db()
    existing = conn.execute('SELECT id FROM holidays WHERE holiday_date = ?', (holiday_date,)).fetchone()
    if existing:
        flash('A holiday is already set for that date. Delete it first if you want to rename it.', 'error')
        return redirect(url_for('holidays'))

    conn.execute(
        'INSERT INTO holidays (holiday_date, name, created_at) VALUES (?, ?, ?)',
        (holiday_date, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    )
    conn.commit()
    flash('Holiday added successfully.', 'success')
    return redirect(url_for('holidays'))


@app.route('/holidays/<int:holiday_id>/delete', methods=['POST'])
@roles_required('admin')
def delete_holiday(holiday_id):
    conn = db.get_db()
    conn.execute('DELETE FROM holidays WHERE id = ?', (holiday_id,))
    conn.commit()
    flash('Holiday deleted.', 'success')
    return redirect(url_for('holidays'))


# ---------------------------------------------------------------------------
# Attendance Management & Reports
# ---------------------------------------------------------------------------

@app.route('/report/attendance')
@login_required
def report_attendance():
    conn = db.get_db()
    employees = conn.execute('SELECT * FROM employees ORDER BY name ASC').fetchall()
    
    base_query = '''
        SELECT a.*, e.name AS employee_name, e.employee_id AS custom_emp_id
        FROM attendance a
        LEFT JOIN employees e ON a.employee_id = e.id
    '''
    where_clauses = []
    params = []

    filter_employee_id = None
    search_emp_id = request.args.get('search_emp_id', '').strip()

    if session.get('role') == 'user':
        my_emp = get_current_employee()
        if my_emp:
            where_clauses.append('a.employee_id = ?')
            params.append(my_emp['id'])
        else:
            where_clauses.append('1 = 0')
    else:
        req_emp = request.args.get('employee_id', '').strip()
        if req_emp:
            filter_employee_id = int(req_emp)
            where_clauses.append('a.employee_id = ?')
            params.append(filter_employee_id)

    if search_emp_id:
        where_clauses.append('e.employee_id LIKE ?')
        params.append(f'%{search_emp_id}%')

    if where_clauses:
        base_query += ' WHERE ' + ' AND '.join(where_clauses)

    base_query += ' ORDER BY a.work_date DESC, a.id DESC'

    rows = conn.execute(base_query, params).fetchall()

    return render_template(
        'report_attendance.html',
        rows=rows,
        employees=employees,
        filter_employee_id=filter_employee_id,
        search_emp_id=search_emp_id,
        user_role=session.get('role'),
    )


@app.route('/report/attendance/add', methods=['POST'])
@roles_required('admin', 'editor')
def add_attendance():
    conn = db.get_db()
    employee_id = request.form.get('employee_id', '').strip()
    work_date = request.form.get('work_date', '').strip()
    check_in = request.form.get('check_in', '').strip() or None
    check_out = request.form.get('check_out', '').strip() or None
    status = request.form.get('status', 'Present').strip()

    if not employee_id or not work_date:
        flash('Employee and Work Date are required.', 'error')
        return redirect(url_for('report_attendance'))

    existing = conn.execute(
        'SELECT id FROM attendance WHERE employee_id = ? AND work_date = ?',
        (employee_id, work_date)
    ).fetchone()

    if existing:
        flash('An attendance record for this employee on this date already exists. Edit it instead.', 'error')
        return redirect(url_for('report_attendance'))

    conn.execute(
        'INSERT INTO attendance (employee_id, work_date, check_in, check_out, status) VALUES (?, ?, ?, ?, ?)',
        (employee_id, work_date, check_in, check_out, status)
    )
    conn.commit()
    flash('Attendance record added successfully.', 'success')
    return redirect(url_for('report_attendance'))


@app.route('/attendance/<int:attendance_id>/edit', methods=['GET', 'POST'])
@roles_required('admin', 'editor')
def edit_attendance(attendance_id):
    att = get_attendance_or_404(attendance_id)
    if att is None:
        flash('That attendance record no longer exists.', 'error')
        return redirect(url_for('report_attendance'))

    conn = db.get_db()
    employees = conn.execute('SELECT * FROM employees ORDER BY name ASC').fetchall()

    if request.method == 'POST':
        employee_id = request.form.get('employee_id', '').strip()
        work_date = request.form.get('work_date', '').strip()
        check_in = request.form.get('check_in', '').strip() or None
        check_out = request.form.get('check_out', '').strip() or None
        status = request.form.get('status', 'Present').strip()

        if not employee_id or not work_date:
            flash('Employee and Work Date are required.', 'error')
            return render_template('attendance_edit.html', att=att, employees=employees)

        conn.execute(
            'UPDATE attendance SET employee_id = ?, work_date = ?, check_in = ?, check_out = ?, status = ? WHERE id = ?',
            (employee_id, work_date, check_in, check_out, status, attendance_id)
        )
        conn.commit()
        flash('Attendance record updated.', 'success')
        return redirect(url_for('report_attendance'))

    return render_template('attendance_edit.html', att=att, employees=employees)


@app.route('/attendance/<int:attendance_id>/delete', methods=['POST'])
@roles_required('admin', 'editor')
def delete_attendance(attendance_id):
    conn = db.get_db()
    conn.execute('DELETE FROM attendance WHERE id = ?', (attendance_id,))
    conn.commit()
    flash('Attendance record deleted.', 'success')
    return redirect(url_for('report_attendance'))


# ---------------------------------------------------------------------------
# Reports (Lists & Excel exports)
# ---------------------------------------------------------------------------

@app.route('/report/employees')
@login_required
def report_employees():
    conn = db.get_db()
    employees = conn.execute('SELECT * FROM employees ORDER BY id ASC').fetchall()
    return render_template('report_employees.html', employees=employees, user_role=session.get('role'))


@app.route('/export/employees')
@roles_required('admin', 'editor')
def export_employees():
    conn = db.get_db()
    rows = conn.execute(
        'SELECT id, name, employee_id, designation, dob, date_of_joining, created_at FROM employees ORDER BY id ASC'
    ).fetchall()

    formatted_rows = []
    for r in rows:
        formatted_rows.append([
            r['id'],
            r['name'],
            r['employee_id'],
            r['designation'],
            format_date_ddmmyyyy(r['dob']),
            format_date_ddmmyyyy(r['date_of_joining']),
            format_datetime_ddmmyyyy(r['created_at']),
        ])

    headers = ['ID', 'Name', 'Employee ID', 'Designation', 'Date of Birth', 'Date of Joining', 'Created At']
    return make_xlsx_response(headers, formatted_rows, 'employees_report.xlsx')


@app.route('/report/leave')
@login_required
def report_leave():
    conn = db.get_db()
    if session.get('role') == 'user':
        rows = conn.execute(
            'SELECT * FROM leave_requests WHERE submitted_by = ? ORDER BY id DESC',
            (session.get('username'),)
        ).fetchall()
    else:
        rows = conn.execute('SELECT * FROM leave_requests ORDER BY id DESC').fetchall()
    return render_template('report_leave.html', rows=rows, user_role=session.get('role'))


@app.route('/export/leave')
@roles_required('admin', 'editor')
def export_leave():
    conn = db.get_db()
    rows = conn.execute(
        'SELECT id, employee_id, dates, num_days, reason, description, request_date, status, submitted_by, created_at '
        'FROM leave_requests ORDER BY id ASC'
    ).fetchall()

    formatted_rows = []
    for r in rows:
        formatted_rows.append([
            r['id'],
            r['employee_id'],
            r['dates'],
            r['num_days'],
            r['reason'],
            r['description'],
            format_date_ddmmyyyy(r['request_date']),
            r['status'],
            r['submitted_by'],
            format_datetime_ddmmyyyy(r['created_at']),
        ])

    headers = [
        'ID', 'Employee ID', 'Dates', 'Num Days', 'Reason', 'Description',
        'Request Date', 'Status', 'Submitted By', 'Created At'
    ]
    return make_xlsx_response(headers, formatted_rows, 'leave_requests_report.xlsx')


@app.route('/export/attendance')
@roles_required('admin', 'editor')
def export_attendance():
    conn = db.get_db()
    rows = conn.execute(
        '''
        SELECT a.id, e.employee_id AS custom_emp_id, e.name AS employee_name,
               a.work_date, a.check_in, a.check_out, a.status
        FROM attendance a
        LEFT JOIN employees e ON a.employee_id = e.id
        ORDER BY a.work_date DESC, a.id DESC
        '''
    ).fetchall()

    formatted_rows = []
    for r in rows:
        formatted_rows.append([
            r['id'],
            r['custom_emp_id'],
            r['employee_name'],
            format_date_ddmmyyyy(r['work_date']),
            r['check_in'],
            r['check_out'],
            r['status'],
        ])

    headers = ['ID', 'Employee ID', 'Employee Name', 'Work Date', 'Check In', 'Check Out', 'Status']
    return make_xlsx_response(headers, formatted_rows, 'attendance_report.xlsx')


# ---------------------------------------------------------------------------
# Email Feature Route
# ---------------------------------------------------------------------------

@app.route('/send-attendance-email', methods=['POST'])
@login_required
def send_attendance_email():
    sender_email = os.environ.get('MAIL_SENDER_EMAIL')
    sender_password = os.environ.get('MAIL_SENDER_PASSWORD')
    recipient_email = os.environ.get('MAIL_RECIPIENT_EMAIL')

    if not sender_email or not sender_password or not recipient_email:
        flash('Email settings are missing in environment variables.', 'error')
        return redirect(url_for('report_attendance'))

    subject = "Attendance Report"
    body = "Hello,\n\nPlease find your requested attendance report attached.\n\nBest Regards,\nHR System"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        flash('Attendance report sent to your email successfully!', 'success')
    except Exception as e:
        flash(f'Failed to send email: {e}', 'error')

    return redirect(url_for('report_attendance'))


# ---------------------------------------------------------------------------
# Main Entry Point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    app.run(debug=True, port=5000)
