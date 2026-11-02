import sqlite3
import os
import sys
import socket
import webbrowser
import threading
import ctypes
from datetime import datetime
from io import BytesIO
from flask import Flask, request, redirect, url_for, render_template_string, send_file, make_response, jsonify, send_from_directory
from openpyxl import Workbook
import json
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'journal.db')
app = Flask(__name__)
# Скрываем консольное окно
if sys.platform == 'win32':
    whnd = ctypes.windll.kernel32.GetConsoleWindow()
    if whnd != 0:
        ctypes.windll.user32.ShowWindow(whnd, 0)
# Ваш список исполнителей
EXECUTORS_LIST = [
    "Котельников. С. А.",
    "Ишеев Б. М.",
    "Батомункуев Ц. Э.",
    "Абилов М. М.",
    "Буринов И. Ю.",
    "Галяутдинов Э. И.",
    "Дианов М. А.",
    "Кудреватов П. В.",
    "Драгун И. И."
]
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT, area TEXT, created_at TEXT NOT NULL, 
        project_no TEXT, horizon TEXT, work_type TEXT, created_by TEXT NOT NULL,
        phone TEXT, safety TEXT, priority TEXT, executor TEXT, completion_date TEXT, 
        status TEXT, comment TEXT, updated_at TEXT)''')
    
    c.execute('''CREATE TRIGGER IF NOT EXISTS prevent_created_at_update BEFORE UPDATE ON requests
        FOR EACH ROW WHEN OLD.created_at <> NEW.created_at BEGIN SELECT RAISE(ABORT, 'immutable'); END''')
    c.execute('''CREATE TRIGGER IF NOT EXISTS prevent_delete BEFORE DELETE ON requests
        FOR EACH ROW BEGIN SELECT RAISE(ABORT, 'delete forbidden'); END''')
    conn.commit()
    conn.close()
def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try: s.connect(("8.8.8.8", 80)); return s.getsockname()[0]
    except: return "127.0.0.1"
    finally: s.close()
# --- ФУНКЦИЯ УМНОЙ СОРТИРОВКИ ---
def get_sorted_requests():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM requests").fetchall()
    conn.close()
    today_str = datetime.now().strftime("%d.%m.%Y")
    def sort_key(row):
        # Проверяем, подана ли заявка сегодня
        is_today = row['created_at'].startswith(today_str)
        # Проверяем, активна ли заявка (не выполнена и не отклонена)
        is_active = row['status'] not in ['Выполнена', 'Отклонена']
        # Распределяем по группам (чем меньше цифра, тем выше в списке)
        if is_today:
            group = 1  # Сегодняшние заявки (любые)
        elif is_active:
            group = 2  # Старые, но еще не выполненные (не забываем про них)
        else:
            group = 3  # Старый архив (выполненные/отклоненные)
        # Внутри каждой группы сортируем по ID (сначала новые)
        return (group, -row['id'])
    return sorted(rows, key=sort_key)
HTML_PAGE = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Журнал заявок</title>
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#222;padding:20px}
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);max-width:100%;margin:0 auto}
table{width:100%;border-collapse:collapse;margin-top:15px;font-size:11px;}
th,td{padding:6px;border-bottom:1px solid #eee;text-align:left;word-wrap:break-word; vertical-align:top;}
th{background:#f8f9fa; font-weight:bold;}
.header-container { display: flex; align-items: center; gap: 15px; margin-bottom: 20px; }
.header-logo { height: 50px; max-width: 150px; object-fit: contain; }
.header-title { margin: 0; font-size: 24px; }
input,select,textarea{width:100%;padding:6px;margin:3px 0;box-sizing:border-box;font-family:inherit;font-size:12px; border: 1px solid #ddd; border-radius:4px;}
input:focus, select:focus { border-color: #1e3a5f; outline: none; }
.btn{background:#1e3a5f;color:#fff;padding:12px;border:none;border-radius:4px;cursor:pointer;width:100%;margin-top:15px;font-weight:bold; font-size:14px;}
.btn:hover { background: #16304e; }
.col-form{display:flex;gap:15px;margin-bottom:10px;}
.col{flex:1;} label{font-size:11px;color:#555;font-weight:bold; display:block; margin-top:5px;} a{color:#1e3a5f;}
td select, td input[type="date"], td textarea.comment-input {
    width: 100%; padding: 4px; font-size: 11px; margin: 0; border: 1px solid transparent;
    background: transparent; color: inherit; box-sizing: border-box; font-family: inherit;
    border-radius: 3px; transition: all 0.2s; appearance: none; -webkit-appearance: none; -moz-appearance: none;
}
td select {
    background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%23999%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
    background-repeat: no-repeat; background-position: right 4px center; background-size: 8px;
    padding-right: 15px; background-color: transparent; opacity: 0.7;
}
td select:hover, td input[type="date"]:hover, td textarea.comment-input:hover {
    border-color: #ccc; background-color: #fff; opacity: 1;
}
td select:focus, td input[type="date"]:focus, td textarea.comment-input:focus {
    border-color: #1e3a5f; background-color: #fff; outline: none; opacity: 1; box-shadow: 0 0 3px rgba(30,58,95,0.2);
}
textarea.comment-input { resize: vertical; min-height: 26px; overflow: hidden; }
.filter-row th { padding: 4px 2px; background: #fafafa; border-bottom: 2px solid #e9ecef;}
.col-filter { 
    width: 100%; padding: 2px 4px; font-size: 10px; border: 1px solid transparent; 
    background: transparent; border-radius: 3px; color: #555;
    appearance: none; -webkit-appearance: none;
    background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%23ccc%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
    background-repeat: no-repeat; background-position: right center; background-size: 6px;
}
.col-filter:hover, .col-filter:focus { border-color: #ccc; background-color: #fff; color: #000; }
/* Цвета статусов */
.high { background-color:#ffe6e6; }
.high td { color: #b03030; }
.done { background-color:#e6ffe6; }
.done td, .done input, .done select, .done textarea { color: #000 !important; }
.rejected { background-color: #f2f2f2; opacity: 0.75; }
.rejected td, .rejected input, .rejected select, .rejected textarea { 
    text-decoration: line-through !important; color: #777 !important; 
}
.sync-badge { font-size: 11px; color: #2a6e2a; display: flex; align-items: center; gap: 5px; margin-bottom:10px;}
.dot { height: 8px; width: 8px; background-color: #2a6e2a; border-radius: 50%; display: inline-block; animation: blink 2s infinite; }
@keyframes blink { 0% { opacity: 1; } 50% { opacity: 0.3; } 100% { opacity: 1; } }
</style></head><body><div class="card">
<div class="header-container">
    <img src="/logo.png" class="header-logo" alt="Логотип" onerror="this.style.display='none'">
    <h2 class="header-title">Журнал заявок маркшейдерской службы</h2>
</div>
<form method="POST" action="/create">
<div class="col-form">
    <div class="col">
        <label>Участок:</label><input name="area" value="">
        <label>№ проекта:</label><input name="project_no" value="">
        <label>Горизонт:</label><input name="horizon" value="">
        <label>Вид работ:</label><textarea name="work_type" rows="2"></textarea>
    </div>
    <div class="col">
        <label>ФИО лица, подавшего заявку:</label><input name="author" required value="{{ saved_author }}">
        <label>Контактный телефон:</label><input name="phone" value="{{ saved_phone }}">
        <label>Безопасно / Нет / Не требуется:</label>
        <select name="safety">
            <option value="Безопасно">Безопасно</option>
            <option value="Нет">Нет (опасно)</option>
            <option value="Не требуется">Не требуется</option>
        </select>
        <label>Приоритет:</label>
        <select name="priority">
            <option value="Низкий">Низкий</option>
            <option value="Обычный" selected>Обычный</option>
            <option value="Высокий">Высокий</option>
            <option value="Срочный">Срочный</option>
        </select>
    </div>
    <div class="col">
        <label>Примечание (при подаче):</label><textarea name="comment" rows="4"></textarea>
        <label>Дата и время подачи (не меняется):</label><input value="{{ now }}" readonly style="background:#eee;">
    </div>
</div>
<button type="submit" class="btn">ПОДАТЬ ЗАЯВКУ</button>
</form>
<hr style="margin:20px 0; border:0; border-top:1px solid #ddd;">
<div style="display:flex; justify-content:space-between; align-items:center;">
    <div class="sync-badge"><span class="dot"></span> Данные обновляются автоматически</div>
    <a href="/export" style="color: #888; font-size: 12px; text-decoration: none; border-bottom: 1px dashed #888;" onmouseover="this.style.color='#333'" onmouseout="this.style.color='#888'">Выгрузка в Excel (Р7)</a>
</div>
<table id="main-table">
<thead>
<tr>
    <th>№</th><th>Участок</th><th>Дата и время</th><th>№ проекта</th><th>Горизонт</th><th>Вид работ</th><th>ФИО подавшего</th><th>Телефон</th><th>Безопасность</th><th>Приоритет</th>
    <th style="background:#e8f4e8;">Исполнитель</th>
    <th style="background:#e8f4e8;">Дата вып.</th>
    <th style="background:#e8f4e8;">Статус</th>
    <th style="background:#e8f4e8; min-width: 200px;">Примечание</th>
</tr>
<tr class="filter-row">
    <th></th>
    <th><select class="col-filter" data-col="1" onchange="applyFilters()"><option value="">Все</option></select></th>
    <th></th>
    <th></th>
    <th><select class="col-filter" data-col="4" onchange="applyFilters()"><option value="">Все</option></select></th>
    <th></th>
    <th></th>
    <th></th>
    <th><select class="col-filter" data-col="8" onchange="applyFilters()"><option value="">Все</option></select></th>
    <th><select class="col-filter" data-col="9" onchange="applyFilters()"><option value="">Все</option></select></th>
    <th><select class="col-filter" data-col="10" onchange="applyFilters()"><option value="">Все</option></select></th>
    <th></th>
    <th><select class="col-filter" data-col="12" onchange="applyFilters()"><option value="">Все</option></select></th>
    <th></th>
</tr>
</thead>
<tbody id="table-body">
{% for req in requests %}
{% set row_class = 'done' if req['status'] == 'Выполнена' else ('rejected' if req['status'] == 'Отклонена' else ('high' if req['priority'] in ['Высокий','Срочный'] else '')) %}
<tr id="row-{{ req['id'] }}" class="request-row {{ row_class }}">
    <td>{{req['id']}}</td>
    <td>{{req['area']}}</td>
    <td>{{req['created_at']}}</td>
    <td>{{req['project_no']}}</td>
    <td>{{req['horizon']}}</td>
    <td>{{req['work_type']}}</td>
    <td>{{req['created_by']}}</td>
    <td>{{req['phone']}}</td>
    <td>{{req['safety']}}</td>
    <td>{{req['priority']}}</td>
    <td>
        <select onchange="updateInline({{req['id']}}, 'executor', this.value)">
            <option value="">-- Выбрать --</option>
            {% for exec_name in executors %}
                <option value="{{ exec_name }}" {% if req['executor'] == exec_name %}selected{% endif %}>{{ exec_name }}</option>
            {% endfor %}
            {% if req['executor'] and req['executor'] not in executors %}
                <option value="{{ req['executor'] }}" selected>{{ req['executor'] }}</option>
            {% endif %}
        </select>
    </td>
    <td>
        <input type="date" class="date-input" value="{{req['completion_date'] or ''}}" 
               onchange="updateInline({{req['id']}}, 'completion_date', this.value)">
    </td>
    <td>
        <select onchange="handleStatusChange({{req['id']}}, this)">
            <option value="Новая" {% if req['status'] == 'Новая' %}selected{% endif %}>Новая</option>
            <option value="Принята" {% if req['status'] == 'Принята' %}selected{% endif %}>Принята</option>
            <option value="В работе" {% if req['status'] == 'В работе' %}selected{% endif %}>В работе</option>
            <option value="Выполнена" {% if req['status'] == 'Выполнена' %}selected{% endif %}>Выполнена</option>
            <option value="Отклонена" {% if req['status'] == 'Отклонена' %}selected{% endif %}>Отклонена</option>
        </select>
    </td>
    <td>
        <textarea rows="1" class="comment-input" onblur="updateInline({{req['id']}}, 'comment', this.value)" oninput="this.style.height = '';this.style.height = this.scrollHeight + 'px'">{{req['comment'] or ''}}</textarea>
    </td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
<script>
    function populateFilters() {
        const rows = document.querySelectorAll('.request-row');
        const filters = {1: new Set(), 4: new Set(), 8: new Set(), 9: new Set(), 10: new Set(), 12: new Set()};
        rows.forEach(row => {
            if (row.cells[1].innerText.trim()) filters[1].add(row.cells[1].innerText.trim());
            if (row.cells[4].innerText.trim()) filters[4].add(row.cells[4].innerText.trim());
            if (row.cells[8].innerText.trim()) filters[8].add(row.cells[8].innerText.trim());
            if (row.cells[9].innerText.trim()) filters[9].add(row.cells[9].innerText.trim());
            
            const exSelect = row.cells[10].querySelector('select');
            if (exSelect && exSelect.value) filters[10].add(exSelect.value);
            
            const stSelect = row.cells[12].querySelector('select');
            if (stSelect && stSelect.value) filters[12].add(stSelect.value);
        });
        for (let col in filters) {
            const select = document.querySelector(`.col-filter[data-col="${col}"]`);
            if (!select) continue;
            const currentVal = select.value;
            select.innerHTML = '<option value="">Все</option>';
            
            const sortedValues = Array.from(filters[col]).sort();
            sortedValues.forEach(val => {
                const opt = document.createElement('option');
                opt.value = val; opt.textContent = val;
                select.appendChild(opt);
            });
            if (sortedValues.includes(currentVal)) select.value = currentVal;
        }
    }
    function applyFilters() {
        const rows = document.querySelectorAll('.request-row');
        const selects = document.querySelectorAll('.col-filter');
        rows.forEach(row => {
            let show = true;
            selects.forEach(select => {
                const colIdx = parseInt(select.getAttribute('data-col'));
                const filterVal = select.value;
                if (filterVal === "") return;
                let cellVal = "";
                const cell = row.cells[colIdx];
                const cellSelect = cell.querySelector('select');
                
                if (cellSelect) {
                    cellVal = cellSelect.value;
                } else {
                    cellVal = cell.innerText.trim();
                }
                if (cellVal !== filterVal) show = false;
            });
            row.style.display = show ? '' : 'none';
        });
    }
    function updateInline(id, field, value) {
        fetch('/update_inline/' + id, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ [field]: value })
        }).then(res => res.json()).then(data => {
            if (data.ok) {
                if (field === 'status') {
                    const row = document.getElementById('row-' + id);
                    if (value === 'Выполнена') {
                        row.className = 'request-row done';
                    } else if (value === 'Отклонена') {
                        row.className = 'request-row rejected';
                    } else if (row.cells[9].innerText.includes('Высокий') || row.cells[9].innerText.includes('Срочный')) {
                        row.className = 'request-row high';
                    } else {
                        row.className = 'request-row';
                    }
                }
            }
        });
    }
    function handleStatusChange(id, selectElement) {
        const newStatus = selectElement.value;
        const row = document.getElementById('row-' + id);
        const dateInput = row.querySelector('.date-input');
        
        if (newStatus === 'Выполнена' && !dateInput.value) {
            const today = new Date().toISOString().split('T')[0];
            dateInput.value = today;
            updateInline(id, 'status', newStatus);
            updateInline(id, 'completion_date', today);
        } else {
            updateInline(id, 'status', newStatus);
        }
    }
    function refreshTable() {
        const active = document.activeElement;
        if (active && (active.tagName === 'INPUT' || active.tagName === 'SELECT' || active.tagName === 'TEXTAREA')) {
            return; 
        }
        fetch('/api/requests').then(res => res.json()).then(data => {
            let html = '';
            data.forEach(req => {
                let rowClass = '';
                if (req.status === 'Выполнена') rowClass = 'done';
                else if (req.status === 'Отклонена') rowClass = 'rejected';
                else if ((req.priority === 'Высокий' || req.priority === 'Срочный')) rowClass = 'high';
                
                let execOptions = '<option value="">-- Выбрать --</option>';
                const executorsList = {{ executors_json | safe }};
                executorsList.forEach(name => {
                    execOptions += `<option value="${name}" ${req.executor === name ? 'selected' : ''}>${name}</option>`;
                });
                if (req.executor && !executorsList.includes(req.executor)) {
                    execOptions += `<option value="${req.executor}" selected>${req.executor}</option>`;
                }
                html += `<tr id="row-${req.id}" class="request-row ${rowClass}">
                    <td>${req.id}</td><td>${req.area || ''}</td><td>${req.created_at}</td>
                    <td>${req.project_no || ''}</td><td>${req.horizon || ''}</td><td>${req.work_type || ''}</td>
                    <td>${req.created_by}</td><td>${req.phone || ''}</td><td>${req.safety || ''}</td><td>${req.priority}</td>
                    
                    <td><select onchange="updateInline(${req.id}, 'executor', this.value)">${execOptions}</select></td>
                    <td><input type="date" class="date-input" value="${req.completion_date || ''}" onchange="updateInline(${req.id}, 'completion_date', this.value)"></td>
                    <td>
                        <select onchange="handleStatusChange(${req.id}, this)">
                            <option value="Новая" ${req.status === 'Новая' ? 'selected' : ''}>Новая</option>
                            <option value="Принята" ${req.status === 'Принята' ? 'selected' : ''}>Принята</option>
                            <option value="В работе" ${req.status === 'В работе' ? 'selected' : ''}>В работе</option>
                            <option value="Выполнена" ${req.status === 'Выполнена' ? 'selected' : ''}>Выполнена</option>
                            <option value="Отклонена" ${req.status === 'Отклонена' ? 'selected' : ''}>Отклонена</option>
                        </select>
                    </td>
                    <td><textarea rows="1" class="comment-input" onblur="updateInline(${req.id}, 'comment', this.value)" oninput="this.style.height = '';this.style.height = this.scrollHeight + 'px'">${req.comment || ''}</textarea></td>
                </tr>`;
            });
            document.getElementById('table-body').innerHTML = html;
            
            document.querySelectorAll('.comment-input').forEach(ta => {
                ta.style.height = '';
                ta.style.height = ta.scrollHeight + 'px';
            });
            populateFilters();
            applyFilters();
        }).catch(e => console.log(e));
    }
    
    populateFilters();
    setInterval(refreshTable, 2000);
</script>
</body></html>
"""
@app.route('/logo.png')
def logo():
    logo_path = os.path.join(BASE_DIR, 'logo.png')
    if os.path.exists(logo_path):
        return send_from_directory(BASE_DIR, 'logo.png')
    else:
        return "", 404
@app.route('/')
def index():
    rows = get_sorted_requests()
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    saved_author = request.cookies.get('saved_author', '')
    saved_phone = request.cookies.get('saved_phone', '')
    
    return render_template_string(HTML_PAGE, 
                                  requests=rows, 
                                  now=now, 
                                  saved_author=saved_author, 
                                  saved_phone=saved_phone,
                                  executors=EXECUTORS_LIST,
                                  executors_json=json.dumps(EXECUTORS_LIST))
@app.route('/api/requests')
def api_requests():
    # Возвращаем данные уже отсортированными
    rows = get_sorted_requests()
    return jsonify([dict(ix) for ix in rows])
@app.route('/create', methods=['POST'])
def create():
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""INSERT INTO requests 
        (area, created_at, project_no, horizon, work_type, created_by, phone, safety, priority, executor, completion_date, status, comment, updated_at) 
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (request.form.get('area',''), now, request.form.get('project_no',''), request.form.get('horizon',''), 
         request.form.get('work_type',''), request.form['author'], request.form.get('phone',''), 
         request.form['safety'], request.form['priority'], '', '', 'Новая', request.form.get('comment',''), now))
    conn.commit(); conn.close()
    
    resp = make_response(redirect(url_for('index')))
    resp.set_cookie('saved_author', request.form['author'], max_age=365*24*60*60)
    resp.set_cookie('saved_phone', request.form.get('phone',''), max_age=365*24*60*60)
    return resp
@app.route('/update_inline/<int:id>', methods=['POST'])
def update_inline(id):
    data = request.get_json()
    if not data: return jsonify({"ok": False}), 400
    
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    field = list(data.keys())[0]
    value = list(data.values())[0]
    
    if field in ['executor', 'completion_date', 'status', 'comment']:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(f"UPDATE requests SET {field}=?, updated_at=? WHERE id=?", (value, now, id))
        conn.commit(); conn.close()
        return jsonify({"ok": True})
    
    return jsonify({"ok": False, "error": "Invalid field"}), 400
@app.route('/export')
def export():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("""SELECT id, area, created_at, project_no, horizon, work_type, created_by, phone, safety, priority, executor, completion_date, status, comment 
                           FROM requests ORDER BY id""").fetchall()
    conn.close()
    wb = Workbook(); ws = wb.active
    ws.append(["№", "Участок", "Дата и время", "№ проекта", "Горизонт", "Вид работ", "ФИО подавшего", "Телефон", "Безопасность", "Приоритет", "Исполнитель", "Дата вып.", "Статус", "Примечание"])
    for r in rows: ws.append(list(r))
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Журнал.xlsx")
def open_browser():
    import time; time.sleep(1.5); webbrowser.open("http://127.0.0.1:5001")
if __name__ == '__main__':
    init_db()
    ip = get_local_ip()
    print(f"СЕРВЕР РАБОТАЕТ. Адрес: http://{ip}:5001")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='0.0.0.0', port=5001, debug=False, use_reloader=False)
