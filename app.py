import sqlite3
import os
import sys
import socket
import webbrowser
import threading
import ctypes
from datetime import datetime
from io import BytesIO
from flask import Flask, request, redirect, url_for, render_template_string, send_file, make_response, jsonify
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
HTML_PAGE = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Журнал заявок</title>
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#222;padding:20px}
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);max-width:100%;margin:0 auto}
table{width:100%;border-collapse:collapse;margin-top:15px;font-size:11px;}
th,td{padding:6px;border-bottom:1px solid #eee;text-align:left;word-wrap:break-word; vertical-align:top;}
th{background:#f8f9fa; font-weight:bold;}
/* Стили для фильтров в таблице */
.filter-row th { padding: 2px; background: #e9ecef; }
.col-filter { width: 100%; padding: 2px; font-size: 10px; border: 1px solid #ccc; border-radius: 2px;}
input,select,textarea{width:100%;padding:6px;margin:3px 0;box-sizing:border-box;font-family:inherit;font-size:12px; border: 1px solid #ddd;}
input:focus, select:focus { border-color: #1e3a5f; outline: none; }
td input, td select { padding: 4px; font-size: 11px; margin: 0; border: 1px solid #ccc; background: #fff;}
.btn{background:#1e3a5f;color:#fff;padding:12px;border:none;border-radius:4px;cursor:pointer;width:100%;margin-top:15px;font-weight:bold; font-size:14px;}
.btn:hover { background: #16304e; }
.col-form{display:flex;gap:15px;margin-bottom:10px;}
.col{flex:1;} label{font-size:11px;color:#555;font-weight:bold; display:block; margin-top:5px;} a{color:#1e3a5f;}
/* Цвета статусов */
.high { background-color:#ffe6e6; color:#b03030; font-weight:bold; } /* Красный для срочных */
.done { background-color:#e6ffe6; color:#2a6e2a; } /* Зеленый для выполненных */
.sync-badge { font-size: 11px; color: #2a6e2a; display: flex; align-items: center; gap: 5px; margin-bottom:10px;}
.dot { height: 8px; width: 8px; background-color: #2a6e2a; border-radius: 50%; display: inline-block; animation: blink 2s infinite; }
@keyframes blink { 0% { opacity: 1; } 50% { opacity: 0.3; } 100% { opacity: 1; } }
</style></head><body><div class="card">
<h2 style="margin-top:0;">Журнал заявок маркшейдерской службы</h2>
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
    <th style="background:#e8f4e8;">Примечание</th>
</tr>
<!-- Строка с фильтрами (как в Excel) -->
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
{% set row_class = 'done' if req['status'] == 'Выполнена' else ('high' if req['priority'] in ['Высокий','Срочный'] and req['status'] != 'Отклонена' else '') %}
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
        <input type="text" value="{{req['comment'] or ''}}" placeholder="Текст..." 
               onblur="updateInline({{req['id']}}, 'comment', this.value)">
    </td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
<script>
    // --- ФИЛЬТРЫ ---
    // Функция сбора уникальных значений для выпадающих списков
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
            const currentVal = select.value; // Запоминаем текущий выбор
            select.innerHTML = '<option value="">Все</option>';
            
            const sortedValues = Array.from(filters[col]).sort();
            sortedValues.forEach(val => {
                const opt = document.createElement('option');
                opt.value = val; opt.textContent = val;
                select.appendChild(opt);
            });
            // Восстанавливаем выбор, если он еще существует в базе
            if (sortedValues.includes(currentVal)) select.value = currentVal;
        }
    }
    // Функция применения фильтров
    function applyFilters() {
        const rows = document.querySelectorAll('.request-row');
        const selects = document.querySelectorAll('.col-filter');
        rows.forEach(row => {
            let show = true;
            selects.forEach(select => {
                const colIdx = parseInt(select.getAttribute('data-col'));
                const filterVal = select.value;
                if (filterVal === "") return; // Если выбрано "Все" - не фильтруем
                let cellVal = "";
                const cell = row.cells[colIdx];
                const cellSelect = cell.querySelector('select');
                
                // Если в ячейке выпадающий список, берем его значение, иначе текст
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
    // --- СОХРАНЕНИЕ ИЗМЕНЕНИЙ ---
    function updateInline(id, field, value) {
        fetch('/update_inline/' + id, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ [field]: value })
        }).then(res => res.json()).then(data => {
            if (data.ok) {
                // Мгновенно меняем цвет при смене статуса
                if (field === 'status') {
                    const row = document.getElementById('row-' + id);
                    if (value === 'Выполнена') {
                        row.className = 'request-row done';
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
    // --- АВТООБНОВЛЕНИЕ ---
    function refreshTable() {
        const active = document.activeElement;
        // Не обновляем, если пользователь в этот момент печатает или выбирает в таблице
        if (active && (active.tagName === 'INPUT' || active.tagName === 'SELECT' || active.tagName === 'TEXTAREA')) {
            return; 
        }
        fetch('/api/requests').then(res => res.json()).then(data => {
            let html = '';
            data.forEach(req => {
                let rowClass = '';
                if (req.status === 'Выполнена') rowClass = 'done';
                else if ((req.priority === 'Высокий' || req.priority === 'Срочный') && req.status !== 'Отклонена') rowClass = 'high';
                
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
                    <td><input type="text" value="${req.comment || ''}" placeholder="Текст..." onblur="updateInline(${req.id}, 'comment', this.value)"></td>
                </tr>`;
            });
            document.getElementById('table-body').innerHTML = html;
            
            // После обновления перестраиваем фильтры и применяем их
            populateFilters();
            applyFilters();
        }).catch(e => console.log(e));
    }
    
    // Инициализация при загрузке страницы
    populateFilters();
    setInterval(refreshTable, 2000);
</script>
</body></html>
"""
@app.route('/')
def index():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM requests ORDER BY id DESC").fetchall()
    conn.close()
    
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
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM requests ORDER BY id DESC").fetchall()
    conn.close()
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
