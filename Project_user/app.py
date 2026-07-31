import io
import os
import traceback
import calendar as cal
from datetime import datetime, timedelta
from functools import wraps

# ---------------------------------------------------------------------------
# Timezone
# ---------------------------------------------------------------------------
def get_local_time():
    """Return current time in IST (UTC+5:30) as a datetime object."""
    return datetime.utcnow() + timedelta(hours=5, minutes=30)

from dotenv import load_dotenv
load_dotenv()

from flask import (
    Flask, render_template, request, redirect, url_for, session,
    flash, send_file, abort
)
from openpyxl import Workbook
import db

app = Flask(__name__, static_folder='static', instance_relative_config=True)

app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'crest@#2026')

os.makedirs(app.instance_path, exist_ok=True)
app.config['DATABASE'] = os.path.join(app.instance_path, 'data.db')

db.init_app(app)


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
        return datetime.strptime(value, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %I:%M:%S %p')
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
# Auth routes
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
# Dashboard & Management Routes
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
        today = get_local_time().strftime('%Y-%m-%d')
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

        local_timestamp = get_local_time().strftime('%Y-%m-%d %H:%M:%S')
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
    conn.commit()
    flash('User deleted successfully.', 'success')
    return redirect(url_for('report_users'))


# ---------------------------------------------------------------------------
# Employees & Attendance Core Routes
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

        local_timestamp = get_local_time().strftime('%Y-%m-%d %H:%M:%S')
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
    conn.commit()
    return redirect(request.form.get('next') or url_for('report_employees'))


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
            return render_template('leave.html', employee_id=auto_employee_id)

        local_timestamp = get_local_time().strftime('%Y-%m-%d %H:%M:%S')
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


@app.route('/attendance', methods=['GET'])
@login_required
def attendance_home():
    employee = get_current_employee()
    if employee is None:
        flash('Your account isn\'t linked to an employee record yet.', 'error')
        return redirect(url_for('dashboard'))

    conn = db.get_db()
    today = get_local_time().strftime('%Y-%m-%d')
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
        return {'error': 'no employee record linked'}, 400

    conn = db.get_db()
    today = get_local_time().strftime('%Y-%m-%d')
    now = get_local_time().strftime('%I:%M:%S %p')

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
        return {'error': 'no employee record linked'}, 400

    conn = db.get_db()
    today = get_local_time().strftime('%Y-%m-%d')
    now = get_local_time().strftime('%I:%M:%S %p')

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
    now = get_local_time()
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

    holiday_rows = conn.execute(
        "SELECT holiday_date, name FROM holidays "
        "WHERE strftime('%Y', holiday_date) = ? AND strftime('%m', holiday_date) = ?",
        (str(year), f'{month:02d}')
    ).fetchall()
    holiday_map = {r['holiday_date']: r['name'] for r in holiday_rows}

    cal.setfirstweekday(cal.SUNDAY)
    weeks = []
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
                extra_tag = f'Worked on Holiday ({holiday_name})' if holiday_name else ('Worked on Sunday' if is_sunday else None)
            elif holiday_name:
                status_label = 'Holiday'
                css = 'holiday'
                extra_tag = None
            elif is_sunday:
                status_label = 'Weekend'
                css = 'weekend'
                extra_tag = None
            else:
                status_label = None
                css = ''
                extra_tag = None

            week_cells.append({
                'day': day_num, 'iso': iso, 'status': status_label, 'css': css,
                'holiday_name': holiday_name, 'extra_tag': extra_tag, 'is_sunday': is_sunday,
                'check_in': info['check_in'] if info else None,
                'check_out': info['check_out'] if info else None,
                'is_today': iso == now.strftime('%Y-%m-%d'),
            })
        weeks.append(week_cells)

    return render_template(
        'calendar.html', weeks=weeks, month_name=cal.month_name[month],
        year=year, month=month, employees=employees, viewing_id=viewing_id, viewing_name=viewing_name,
        weekday_labels=['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    )


@app.route('/holidays', methods=['GET'])
@login_required
def holidays():
    conn = db.get_db()
    all_holidays = conn.execute('SELECT * FROM holidays ORDER BY holiday_date ASC').fetchall()
    return render_template('holidays.html', holidays=all_holidays, user_role=session.get('role'))


@app.route('/report/attendance')
@login_required
def report_attendance():
    conn = db.get_db()
    employees = conn.execute('SELECT * FROM employees ORDER BY name ASC').fetchall()
    return render_template('report_attendance.html', employees=employees, user_role=session.get('role'))


@app.route('/report')
@login_required
def report():
    conn = db.get_db()
    employee_count = conn.execute('SELECT COUNT(*) AS c FROM employees').fetchone()['c']
    leave_count = conn.execute('SELECT COUNT(*) AS c FROM leave_requests').fetchone()['c']
    return render_template('report.html', employee_count=employee_count, leave_count=leave_count)


@app.route('/report/employees')
@login_required
def report_employees():
    conn = db.get_db()
    employees = conn.execute('SELECT * FROM employees ORDER BY id ASC').fetchall()
    return render_template('report_employees.html', employees=employees, user_role=session.get('role'))


@app.route('/report/leave')
@login_required
def report_leave():
    conn = db.get_db()
    leave_requests = conn.execute('SELECT * FROM leave_requests ORDER BY id ASC').fetchall()
    return render_template('report_leave.html', leave_requests=leave_requests, user_role=session.get('role'))


def _build_attendance_report(role, username):
    conn = db.get_db()
    base_query = '''
        SELECT a.id, e.name AS employee_name, e.employee_id AS custom_emp_id, a.work_date, a.check_in, a.check_out, a.status, a.created_at
        FROM attendance a
        LEFT JOIN employees e ON a.employee_id = e.id
    '''
    rows = conn.execute(base_query).fetchall()
    headers = ['ID', 'Employee Name', 'Employee ID', 'Work Date', 'Check In', 'Check Out', 'Status', 'Created At']
    data = [
        [idx, r['employee_name'] or 'Unknown', r['custom_emp_id'] or '', format_date_ddmmyyyy(r['work_date']),
         r['check_in'], r['check_out'], r['status'], format_datetime_ddmmyyyy(r['created_at'])]
        for idx, r in enumerate(rows, start=1)
    ]
    return headers, data, 'attendance.xlsx'


def _build_leave_report(role, username):
    conn = db.get_db()
    rows = conn.execute('SELECT * FROM leave_requests ORDER BY id ASC').fetchall()
    headers = ['ID', 'Requested By', 'Employee ID', 'Dates', 'Number of Days', 'Reason',
               'Description', 'Date of Request', 'Status', 'Created At']
    data = [
        [idx, r['submitted_by'], r['employee_id'], r['dates'], r['num_days'], r['reason'],
         r['description'], format_date_ddmmyyyy(r['request_date']), r['status'],
         format_datetime_ddmmyyyy(r['created_at'])]
        for idx, r in enumerate(rows, start=1)
    ]
    return headers, data, 'leave_requests.xlsx'


# ---------------------------------------------------------------------------
# Email Routes (HTTP API / Render-Compatible using Resend / Requests)
# ---------------------------------------------------------------------------
import base64
import requests

def send_email_via_http(recipient_email, subject, body, attachments=None):
    """Sends email securely over HTTPS (port 443) using an HTTP API provider 
    like Resend (https://resend.com), bypassing Render's port blocking limits."""
    api_key = os.environ.get('MAIL_API_KEY') # Get free key from resend.com
    sender_email = os.environ.get('MAIL_SENDER_EMAIL', 'onboarding@resend.dev')

    if not api_key:
        raise ValueError("MAIL_API_KEY environment variable is not configured.")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "from": sender_email,
        "to": [recipient_email],
        "subject": subject,
        "text": body,
        "attachments": []
    }

    if attachments:
        for buffer, filename in attachments:
            encoded_content = base64.b64encode(buffer.read()).decode('utf-8')
            payload["attachments"].append({
                "filename": filename,
                "content": encoded_content
            })

    response = requests.post("https://api.resend.com/emails", json=payload, headers=headers)
    if response.status_code not in (200, 201):
        raise Exception(f"API Error: {response.text}")
    return True


@app.route('/send-report-email', methods=['POST'])
@roles_required('admin', 'editor')
def send_report_email():
    recipient_email = os.environ.get('MAIL_RECIPIENT_EMAIL')
    fallback_endpoint = url_for('report')
    redirect_target = request.referrer or fallback_endpoint

    if not recipient_email:
        flash('Recipient email is not configured on the server.', 'error')
        return redirect(redirect_target)

    try:
        role = session.get('role')
        username = session.get('username')

        att_headers, att_data, att_filename = _build_attendance_report(role, username)
        att_buffer = build_xlsx_buffer(att_headers, att_data)

        leave_headers, leave_data, leave_filename = _build_leave_report(role, username)
        leave_buffer = build_xlsx_buffer(leave_headers, leave_data)

        sent_at = get_local_time().strftime('%d-%m-%Y %I:%M:%S %p')
        body = (
            "Hello,\n\n"
            "Please find the attendance report and leave requests report attached.\n\n"
            f"Report generated: {sent_at} (IST)\n\n"
            "Best Regards,\nHR System"
        )

        send_email_via_http(
            recipient_email=recipient_email,
            subject="Attendance & Leave Reports",
            body=body,
            attachments=[(att_buffer, att_filename), (leave_buffer, leave_filename)]
        )

        flash('Attendance and leave reports sent successfully!', 'success')
    except Exception as e:
        app.logger.error("send_report_email failed: %s", e)
        traceback.print_exc()
        flash(f'Failed to send email: {e}', 'error')

    return redirect(redirect_target)


@app.route('/send-attendance-email', methods=['POST'])
@login_required
def send_attendance_email():
    recipient_email = os.environ.get('MAIL_RECIPIENT_EMAIL')
    fallback_endpoint = url_for('report_leave')
    redirect_target = request.referrer or fallback_endpoint

    if not recipient_email:
        flash('Recipient email is not configured on the server.', 'error')
        return redirect(redirect_target)

    try:
        headers, data, filename = _build_attendance_report(session.get('role'), session.get('username'))
        xlsx_buffer = build_xlsx_buffer(headers, data)

        sent_at = get_local_time().strftime('%d-%m-%Y %I:%M:%S %p')
        body = (
            "Hello,\n\n"
            "Please find your requested attendance report attached.\n\n"
            f"Report generated: {sent_at} (IST)\n\n"
            "Best Regards,\nHR System"
        )

        send_email_via_http(
            recipient_email=recipient_email,
            subject="Attendance Report",
            body=body,
            attachments=[(xlsx_buffer, filename)]
        )

        flash('Attendance report sent successfully!', 'success')
    except Exception as e:
        app.logger.error("send_attendance_email failed: %s", e)
        traceback.print_exc()
        flash(f'Failed to send email: {e}', 'error')

    return redirect(redirect_target)


@app.route('/debug-env-check')
@login_required
def debug_env_check():
    return {
        'MAIL_API_KEY_set': bool(os.environ.get('MAIL_API_KEY')),
        'MAIL_RECIPIENT_EMAIL_set': bool(os.environ.get('MAIL_RECIPIENT_EMAIL')),
    }


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5001, debug=True)
