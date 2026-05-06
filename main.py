import os
import sqlite3
import pandas as pd
import re
from datetime import datetime
from functools import wraps
from flask import Flask, request, jsonify, render_template_string, send_from_directory, send_file, Response

app = Flask(__name__)

# Получаем путь к текущей папке
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# Настройки папки для фото чеков
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'receipt_photos')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# ---------------------------------------------------------
# АВТОРИЗАЦИЯ ДЛЯ АДМИНКИ
# ---------------------------------------------------------
ADMIN_LOGIN = "Alko"
ADMIN_PASSWORD = "zxcASDqwe123"

def check_auth(username, password):
    return username == ADMIN_LOGIN and password == ADMIN_PASSWORD

def authenticate():
    # Эта функция вызывает стандартное окно браузера для ввода пароля
    return Response(
        'Требуется авторизация для доступа к панели управления.', 401,
        {'WWW-Authenticate': 'Basic realm="Admin Login Required"'})

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# ---------------------------------------------------------
# АВТОМАТИЧЕСКАЯ ЗАГРУЗКА БАЗЫ МАГАЗИНОВ ИЗ EXCEL
# ---------------------------------------------------------
def load_stores_from_excel(filename="StoreAlk.xlsx"):
    stores = {}
    filepath = os.path.join(BASE_DIR, filename)
    try:
        df = pd.read_excel(filepath)
        for index, row in df.iterrows():
            city = str(row['Регион']).strip()
            address = str(row['Адрес']).strip()
            if pd.isna(row['Регион']) or pd.isna(row['Адрес']):
                continue
            if city not in stores:
                stores[city] = []
            stores[city].append(address)
        print(f"Успешно загружено {len(stores)} городов.")
    except Exception as e:
        print(f"Ошибка загрузки Excel: {e}")
        stores = {"Ошибка": ["Файл StoreAlk.xlsx не найден"]}
    return stores

STORES_DATA = load_stores_from_excel()

def init_db():
    conn = sqlite3.connect(os.path.join(BASE_DIR, 'database.db'))
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS receipts
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  user_name TEXT,
                  user_phone TEXT,
                  city TEXT, 
                  address TEXT, 
                  purchase_date TEXT, 
                  purchase_time TEXT, 
                  receipt_number TEXT, 
                  photo_path TEXT,
                  UNIQUE(city, address, receipt_number))''')
    conn.commit()
    conn.close()

# ---------------------------------------------------------
# СТРАНИЦА ПОКУПАТЕЛЯ (ЧИСТЫЙ ЛЕНДИНГ С ЛОГОТИПОМ)
# ---------------------------------------------------------
HTML_PAGE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Розыгрыш | Алкотека</title>
    
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Merriweather:wght@400;700&display=swap" rel="stylesheet">
    
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

    <style>
        :root {
            --brand-color: #3b0918; 
            --bg-color: #f9f9f9;
            --text-main: #222;
            --text-muted: #888;
            --border-color: #e5e5e5;
        }

        body { 
            font-family: 'Inter', sans-serif; 
            background-color: var(--bg-color); 
            margin: 0; 
            color: var(--text-main);
        }

        .site-header {
            background: #fff;
            border-bottom: 1px solid var(--border-color);
            padding: 15px 20px;
            display: flex;
            justify-content: center; 
            align-items: center;
        }
        
        .site-logo img {
            max-height: 50px; 
            width: auto;
            object-fit: contain;
        }

        .page-content {
            padding: 50px 20px;
            display: flex;
            justify-content: center;
        }

        .card { 
            background: #ffffff; 
            width: 100%;
            max-width: 500px; 
            padding: 40px; 
            border-radius: 16px; 
            box-shadow: 0 10px 40px rgba(0,0,0,0.04);
            border: 1px solid var(--border-color);
        }

        h2 { 
            font-family: 'Merriweather', serif; 
            color: var(--text-main); 
            font-weight: 400; 
            font-size: 28px; 
            margin: 0 0 10px 0; 
            text-align: center;
        }

        .subtitle { 
            color: var(--text-muted); 
            font-size: 14px; 
            text-align: center; 
            margin-bottom: 30px; 
        }

        .form-group { margin-bottom: 20px; }

        label { 
            display: block; 
            margin-bottom: 6px; 
            font-weight: 500; 
            font-size: 13px; 
            color: var(--text-muted); 
        }

        input[type="text"], input[type="tel"], input[type="date"], input[type="time"] { 
            width: 100%; 
            padding: 14px 16px; 
            background: #fff;
            border: 1px solid #ccc; 
            border-radius: 8px; 
            box-sizing: border-box; 
            font-family: 'Inter', sans-serif;
            font-size: 15px;
            transition: border-color 0.2s;
        }

        input:focus { 
            outline: none; 
            border-color: var(--brand-color); 
        }

        .row { display: flex; gap: 15px; }
        .row > div { flex: 1; }

        input[type="file"] {
            width: 100%;
            padding: 10px;
            border: 1px dashed #ccc;
            border-radius: 8px;
            cursor: pointer;
            font-size: 13px;
            color: var(--text-main);
        }

        .select2-container--default .select2-selection--single { 
            height: 48px !important; 
            border: 1px solid #ccc !important; 
            border-radius: 8px !important; 
            display: flex; 
            align-items: center; 
        }
        .select2-container--default .select2-selection--single .select2-selection__rendered {
            font-family: 'Inter', sans-serif;
            font-size: 15px;
            color: var(--text-main);
            padding-left: 16px;
        }
        .select2-container--default .select2-selection--single .select2-selection__arrow { height: 46px !important; right: 10px !important; }
        .select2-container { width: 100% !important; }

        button { 
            background: var(--brand-color); 
            color: white; 
            border: none; 
            padding: 16px; 
            width: 100%; 
            border-radius: 30px; 
            font-size: 16px; 
            font-weight: 500; 
            margin-top: 10px; 
            cursor: pointer; 
            transition: opacity 0.2s; 
        }

        button:hover { opacity: 0.9; }
        button:disabled { background: #ccc; cursor: not-allowed; }

        #msg { 
            margin-top: 20px; 
            text-align: center; 
            font-weight: 500; 
            font-size: 14px;
            padding: 15px;
            border-radius: 8px;
            display: none;
        }
        .error { color: #d32f2f; background: #ffebee; display: block !important; }
        .success { color: #2e7d32; background: #e8f5e9; display: block !important; }
        .hint { font-size: 12px; color: var(--text-muted); margin-top: 5px; }

        @media (max-width: 768px) {
            .site-header { padding: 15px; }
            .card { padding: 30px 20px; border: none; box-shadow: none; border-radius: 0; }
            .page-content { padding: 0; background: #fff;}
        }
    </style>
</head>
<body>

    <div class="site-header">
        <a href="#" class="site-logo"><img src="/logo.png" alt="Алкотека"></a>
    </div>

    <div class="page-content">
        <div class="card">
            <h2>Регистрация чека</h2>
            <p class="subtitle">Участвуйте в розыгрыше призов от Алкотеки</p>
            
            <form id="regForm">
                <div class="form-group">
                    <label>Ваше имя</label>
                    <input type="text" id="u_name" placeholder="Иван Иванов" required>
                </div>

                <div class="form-group">
                    <label>Номер телефона</label>
                    <input type="tel" id="u_phone" placeholder="+79001234567" pattern="\+79[0-9]{9}" maxlength="12" required>
                    <div class="hint">Формат: +79XXXXXXXXX</div>
                </div>

                <div class="form-group">
                    <label>Город покупки</label>
                    <select id="city" required>
                        <option value=""></option>
                        {% for city in sorted_cities %}
                        <option value="{{ city }}">{{ city }}</option>
                        {% endfor %}
                    </select>
                </div>

                <div class="form-group">
                    <label>Адрес магазина</label>
                    <select id="address" disabled required>
                        <option value=""></option>
                    </select>
                </div>

                <div class="row form-group">
                    <div>
                        <label>Дата чека</label>
                        <input type="date" id="p_date" required>
                    </div>
                    <div>
                        <label>Время</label>
                        <input type="time" id="p_time" required>
                    </div>
                </div>

                <div class="form-group">
                    <label>Номер чека (ФД)</label>
                    <input type="text" id="r_num" placeholder="Например: 12345" required>
                </div>

                <div class="form-group">
                    <label>Фотография чека</label>
                    <input type="file" id="photo" accept="image/*" required>
                </div>

                <button type="submit" id="submitBtn">Зарегистрировать</button>
            </form>
            <div id="msg"></div>
        </div>
    </div>

    <script>
        const stores = {{ stores_json|safe }};
        
        $(document).ready(function() {
            $('#city').select2({
                placeholder: "Выберите город",
                width: '100%',
                language: { noResults: () => "Город не найден" }
            });

            $('#address').select2({
                placeholder: "Сначала выберите город",
                width: '100%',
                language: { noResults: () => "Адрес не найден" }
            });

            const dateInput = document.getElementById('p_date');
            const today = new Date().toISOString().split('T')[0];
            dateInput.setAttribute('max', today);
            dateInput.setAttribute('min', '2025-01-01');

            $('#city').on('change', function() {
                const city = $(this).val();
                const addrSelect = $('#address');
                
                addrSelect.empty();
                addrSelect.append('<option value=""></option>');

                if (city) {
                    addrSelect.prop('disabled', false);
                    let sortedAddresses = stores[city].sort();
                    sortedAddresses.forEach(a => {
                        addrSelect.append(new Option(a, a));
                    });
                    addrSelect.select2({ placeholder: "Выберите адрес", width: '100%' });
                } else {
                    addrSelect.prop('disabled', true);
                    addrSelect.select2({ placeholder: "Сначала выберите город", width: '100%' });
                }
                addrSelect.trigger('change');
            });
        });

        document.getElementById('u_phone').addEventListener('input', function (e) {
            this.value = this.value.replace(/[^\+0-9]/g, '');
        });

        document.getElementById('regForm').onsubmit = async (e) => {
            e.preventDefault();
            const msg = document.getElementById('msg');
            const submitBtn = document.getElementById('submitBtn');
            
            const cityVal = $('#city').val();
            const addrVal = $('#address').val();
            if (!cityVal || !addrVal) {
                msg.className = 'error';
                msg.innerHTML = 'Пожалуйста, выберите город и адрес магазина.';
                return;
            }

            msg.className = '';
            msg.innerHTML = "Регистрация чека...";
            msg.style.display = 'block';
            submitBtn.disabled = true;

            const formData = new FormData();
            formData.append('name', document.getElementById('u_name').value);
            formData.append('phone', document.getElementById('u_phone').value);
            formData.append('city', cityVal);
            formData.append('address', addrVal);
            formData.append('date', document.getElementById('p_date').value);
            formData.append('time', document.getElementById('p_time').value);
            formData.append('r_num', document.getElementById('r_num').value);
            formData.append('photo', document.getElementById('photo').files[0]);

            try {
                const res = await fetch('/register', { method: 'POST', body: formData });
                const data = await res.json();
                
                msg.className = data.status === 'success' ? 'success' : 'error';
                msg.innerHTML = data.message;
                
                if(data.status === 'success') {
                    e.target.reset();
                    $('#city').val(null).trigger('change');
                }
            } catch (err) {
                msg.className = 'error';
                msg.innerHTML = 'Произошла ошибка соединения с сервером.';
            } finally {
                submitBtn.disabled = false;
            }
        };
    </script>
</body>
</html>
"""

# ---------------------------------------------------------
# ПАНЕЛЬ АДМИНИСТРАТОРА
# ---------------------------------------------------------
ADMIN_PAGE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Панель управления розыгрышем</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background-color: #f9f9f9;}
        table { width: 100%; border-collapse: collapse; margin: 20px 0; background: white;}
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background-color: #3b0918; color: white; }
        a { color: #3b0918; text-decoration: none; font-weight: bold;}
        .header { display: flex; justify-content: space-between; align-items: center; }
        .btn { background-color: #2e7d32; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;}
        .btn:hover { background-color: #1b5e20; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Зарегистрированные участники</h2>
        <a href="/admin/export" class="btn">Скачать в Excel</a>
    </div>
    <table>
        <tr>
            <th>ID</th><th>Имя</th><th>Телефон</th><th>Город</th>
            <th>Адрес</th><th>Дата и Время</th><th>Чек (ФД)</th><th>Фото</th>
        </tr>
        {% for row in rows %}
        <tr>
            <td>{{ row[0] }}</td><td>{{ row[1] }}</td><td>{{ row[2] }}</td><td>{{ row[3] }}</td>
            <td>{{ row[4] }}</td><td>{{ row[5] }} {{ row[6] }}</td><td>{{ row[7] }}</td>
            <td>
                {% if row[8] %}
                <a href="/{{ row[8] }}" target="_blank">Смотреть</a>
                {% endif %}
            </td>
        </tr>
        {% endfor %}
    </table>
</body>
</html>
"""

# ---------------------------------------------------------
# ЛОГИКА ОБРАБОТКИ (БЭКЕНД)
# ---------------------------------------------------------
@app.route('/')
def index():
    sorted_cities = sorted(STORES_DATA.keys())
    return render_template_string(HTML_PAGE, sorted_cities=sorted_cities, stores_json=STORES_DATA)

@app.route('/logo.png')
def serve_logo():
    return send_from_directory(BASE_DIR, 'logo.png')

@app.route('/register', methods=['POST'])
def register():
    name = request.form.get('name')
    phone = request.form.get('phone')
    city = request.form.get('city')
    address = request.form.get('address')
    r_num = request.form.get('r_num')
    p_date = request.form.get('date')
    p_time = request.form.get('time')
    photo = request.files.get('photo')

    if not phone or not re.match(r'^\+79\d{9}$', phone):
        return jsonify({"status": "error", "message": "Ошибка: Номер телефона должен быть в формате +79XXXXXXXXX (12 символов)."})

    try:
        parsed_date = datetime.strptime(p_date, '%Y-%m-%d')
        current_date = datetime.now()
        if parsed_date.year < 2025:
            return jsonify({"status": "error", "message": "Ошибка: К участию принимаются чеки не старше 2025 года."})
        if parsed_date > current_date:
            return jsonify({"status": "error", "message": "Ошибка: Дата покупки не может быть в будущем."})
    except ValueError:
        return jsonify({"status": "error", "message": "Ошибка: Неверный формат даты."})

    photo_path = ""
    if photo:
        filename = f"{phone}_{r_num}_{photo.filename}"
        photo_path = os.path.join('receipt_photos', filename).replace('\\', '/')
        photo.save(os.path.join(UPLOAD_FOLDER, filename))

    try:
        conn = sqlite3.connect(os.path.join(BASE_DIR, 'database.db'))
        c = conn.cursor()
        c.execute("INSERT INTO receipts (user_name, user_phone, city, address, purchase_date, purchase_time, receipt_number, photo_path) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  (name, phone, city, address, p_date, p_time, r_num, photo_path))
        conn.commit()
        conn.close()
        return jsonify({"status": "success", "message": "Чек успешно зарегистрирован! Удачи в розыгрыше."})
    except sqlite3.IntegrityError:
        return jsonify({"status": "error", "message": "Ошибка: этот чек уже зарегистрирован в данном магазине."})

# ДОБАВЛЕН ДЕКОРАТОР @requires_auth ДЛЯ ЗАЩИТЫ ФОТО
@app.route('/receipt_photos/<filename>')
@requires_auth
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

# ДОБАВЛЕН ДЕКОРАТОР @requires_auth ДЛЯ ЗАЩИТЫ АДМИНКИ
@app.route('/admin')
@requires_auth
def admin_panel():
    conn = sqlite3.connect(os.path.join(BASE_DIR, 'database.db'))
    c = conn.cursor()
    c.execute("SELECT * FROM receipts")
    rows = c.fetchall()
    conn.close()
    return render_template_string(ADMIN_PAGE, rows=rows)

# ДОБАВЛЕН ДЕКОРАТОР @requires_auth ДЛЯ ЗАЩИТЫ ВЫГРУЗКИ БАЗЫ
@app.route('/admin/export')
@requires_auth
def export_excel():
    conn = sqlite3.connect(os.path.join(BASE_DIR, 'database.db'))
    df = pd.read_sql_query("SELECT * FROM receipts", conn)
    conn.close()
    df.columns = ['ID', 'Имя', 'Телефон', 'Город', 'Адрес', 'Дата', 'Время', 'Номер чека', 'Путь к фото']
    export_file = os.path.join(BASE_DIR, 'Участники_Розыгрыша.xlsx')
    df.to_excel(export_file, index=False)
    return send_file(export_file, as_attachment=True)

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=8080)