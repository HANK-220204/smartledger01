from flask import Flask, render_template, request, jsonify, session
from functools import wraps
import sqlite3, csv, io, os, json, urllib.request, re
from datetime import date, datetime
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader

app = Flask(__name__)
app.secret_key = os.environ.get('SMARTLEDGER_SECRET_KEY', 'smartledger-development-secret-change-me')
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
DB = os.path.join(os.path.dirname(__file__), 'smartledger.db')

PROVIDER_DEFAULTS = {
    'demo': {'base_url': '', 'model': 'demo'},
    'ollama': {'base_url': 'http://127.0.0.1:11434/v1/chat/completions', 'model': 'qwen2.5:3b'},
    'deepseek': {'base_url': 'https://api.deepseek.com/chat/completions', 'model': 'deepseek-chat'},
    'doubao': {'base_url': 'https://ark.cn-beijing.volces.com/api/v3/chat/completions', 'model': 'doubao-1-5-pro-32k-250115'},
    'kimi': {'base_url': 'https://api.moonshot.cn/v1/chat/completions', 'model': 'moonshot-v1-8k'},
    'openai': {'base_url': 'https://api.openai.com/v1/chat/completions', 'model': 'gpt-4o-mini'},
}

def resolve_provider_settings(provider, base_url='', model=''):
    provider_name = str(provider or 'demo').lower()
    default = PROVIDER_DEFAULTS.get(provider_name, PROVIDER_DEFAULTS['deepseek'])
    resolved_base_url = (base_url or default['base_url']).strip()
    resolved_model = (model or default['model']).strip()
    return provider_name, resolved_base_url, resolved_model

def build_local_plan(summary, transactions, income, save_rate, goal):
    expenses = [item for item in transactions if item['kind'] == 'expense']
    incomes = [item for item in transactions if item['kind'] == 'income']
    category_totals = summary['支出分类']
    total_expense = summary['已记录支出']
    monthly_expenses = {}
    monthly_incomes = {}
    for item in transactions:
        month = str(item.get('tx_date', ''))[:7] or '未知月份'
        target = monthly_incomes if item['kind'] == 'income' else monthly_expenses
        target[month] = target.get(month, 0) + item['amount']
    recent_months = sorted(set(monthly_expenses) | set(monthly_incomes))[-3:]
    average_expense = total_expense / max(len(monthly_expenses), 1)
    average_income = sum(monthly_incomes.values()) / max(len(monthly_incomes), 1)
    top_categories = sorted(category_totals.items(), key=lambda pair: pair[1], reverse=True)
    top_category = top_categories[0] if top_categories else ('暂无', 0)
    concentration = top_category[1] / total_expense if total_expense else 0
    plan_income = income or average_income
    target_saving = max(plan_income * save_rate / 100, 0)
    spendable = max(plan_income - target_saving, 0)
    observations = []
    if not transactions:
        observations.append('目前还没有足够的历史账目，先用本月实际收入和支出建立基线，暂不建议把预算定得过紧。')
    elif len(monthly_incomes) >= 2 and max(monthly_incomes.values()) - min(monthly_incomes.values()) > average_income * 0.3:
        observations.append('收入月份之间波动较大，建议按近几个月较低收入安排固定预算，把额外收入分配给储蓄或应急金。')
    elif average_income and total_expense / max(sum(monthly_incomes.values()), 1) > 0.85:
        observations.append('历史支出已经接近收入，当前储蓄目标需要分阶段执行，先建立小额自动储蓄，避免月底透支。')
    else:
        observations.append('收入与支出记录相对可控，可以采用自动储蓄加分类限额的方式逐步提高结余。')
    if concentration >= 0.4:
        observations.append(f'{top_category[0]}占支出约{concentration:.0%}，它是最值得优先复盘的分类，不建议平均削减所有分类。')
    if len(recent_months) >= 2:
        first, last = recent_months[0], recent_months[-1]
        if monthly_expenses.get(last, 0) > monthly_expenses.get(first, 0) * 1.2:
            observations.append('最近月份支出较前期明显上升，建议先观察最近两周的新增支出，再调整下月限额。')
    category_plan = []
    for category, amount in top_categories:
        ratio = amount / total_expense if total_expense else 0
        limit = amount / max(len(monthly_expenses), 1) * (0.9 if ratio >= 0.25 else 1.0)
        action = '先降低非必要支出并设置周限额' if ratio >= 0.25 else '保持记录，观察是否持续增长'
        category_plan.append(f'- {category}：历史月均约 ¥{limit:,.2f}，占支出 {ratio:.0%}；{action}。')
    return f'''【基于你的真实账目的个性化规划】

一、数据诊断
本次分析了 {len(transactions)} 笔账目，覆盖 {len(recent_months)} 个有记录月份。
计划月收入：¥{plan_income:,.2f}；历史月均收入：¥{average_income:,.2f}；历史月均支出：¥{average_expense:,.2f}。
目标储蓄：¥{target_saving:,.2f}；扣除储蓄后可分配支出：¥{spendable:,.2f}。
{' '.join(observations)}

二、下月预算建议
- 先在收入到账后转出 ¥{target_saving:,.2f}，如果现金流紧张，先执行目标的 50%，连续两个月稳定后再提高。
- 固定支出上限建议为 ¥{spendable * 0.55:,.2f}，可变支出上限建议为 ¥{spendable * 0.35:,.2f}，余下 ¥{spendable * 0.10:,.2f} 作为机动金。
- 预算不是平均分配：优先处理占比最高的分类，其他分类只做记录和提醒。

三、分类行动清单
{chr(10).join(category_plan) or '- 暂无分类数据，请先持续记录。'}

四、执行安排
1. 每周固定一天检查本周累计支出，重点查看 {top_category[0]} 是否超过周限额。
2. 当某分类达到月度限额的 80% 时，暂停非必要消费并重新评估目标。
3. 月末比较预算与实际，不追求一次调整到位，只对连续两个月偏离的分类做 10% 调整。
4. 应急金优先积累到 3 个月必要支出，再考虑提高长期储蓄目标。

规划目标：{goal}
以上内容根据当前账本生成，不构成投资建议。'''

def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    with db() as con:
        con.execute('''CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            phone TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )''')
        con.execute('''CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            kind TEXT NOT NULL CHECK(kind IN ('income','expense')),
            amount REAL NOT NULL,
            category TEXT NOT NULL,
            note TEXT DEFAULT '',
            tx_date TEXT NOT NULL,
            created_at TEXT NOT NULL
        )''')
        columns = {row['name'] for row in con.execute('PRAGMA table_info(transactions)')}
        if 'user_id' not in columns:
            con.execute('ALTER TABLE transactions ADD COLUMN user_id INTEGER')
        if 'infer_category' in globals():
            uncategorized = con.execute("SELECT id, kind, category, note FROM transactions WHERE category IS NULL OR category='' OR category='其他'").fetchall()
            for transaction in uncategorized:
                category = normalize_category(transaction['category'], transaction['note'], transaction['kind'])
                con.execute('UPDATE transactions SET category=? WHERE id=?', (category, transaction['id']))

def current_user_id():
    return session.get('user_id')

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user_id():
            return jsonify(error='请先登录'), 401
        return view(*args, **kwargs)
    return wrapped

def rows():
    with db() as con:
        transactions = [dict(r) for r in con.execute('SELECT * FROM transactions WHERE user_id=? ORDER BY tx_date DESC, id DESC', (current_user_id(),))]
        for transaction in transactions:
            category = normalize_category(transaction.get('category', ''), transaction.get('note', ''), transaction.get('kind', 'expense'))
            if category != transaction['category']:
                transaction['category'] = category
                con.execute('UPDATE transactions SET category=? WHERE id=? AND user_id=?', (category, transaction['id'], current_user_id()))
        return transactions

@app.route('/')
def index():
    if not current_user_id():
        return render_template('auth.html')
    return render_template('index.html', user_email=session.get('email', ''))

@app.post('/api/auth/register')
def register():
    data = request.get_json(force=True)
    email = str(data.get('email', '')).strip().lower()
    phone = str(data.get('phone', '')).strip()
    password = str(data.get('password', ''))
    if not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email):
        return jsonify(error='请输入有效的邮箱地址'), 400
    if not re.fullmatch(r'1\d{10}', phone):
        return jsonify(error='请输入有效的11位手机号'), 400
    if len(password) < 6:
        return jsonify(error='密码至少需要6位'), 400
    try:
        with db() as con:
            cur = con.execute('INSERT INTO users(email,phone,password_hash,created_at) VALUES(?,?,?,?)',
                (email, phone, generate_password_hash(password), datetime.now().isoformat()))
            session['user_id'] = cur.lastrowid
            session['email'] = email
    except sqlite3.IntegrityError:
        return jsonify(error='邮箱或手机号已注册'), 409
    return jsonify(ok=True)

@app.post('/api/auth/login')
def login():
    data = request.get_json(force=True)
    email = str(data.get('email', '')).strip().lower()
    password = str(data.get('password', ''))
    with db() as con:
        user = con.execute('SELECT * FROM users WHERE email=?', (email,)).fetchone()
    if not user or not check_password_hash(user['password_hash'], password):
        return jsonify(error='邮箱或密码不正确'), 401
    session['user_id'] = user['id']
    session['email'] = user['email']
    return jsonify(ok=True)

@app.post('/api/auth/reset')
def reset_password():
    data = request.get_json(force=True)
    email = str(data.get('email', '')).strip().lower()
    phone = str(data.get('phone', '')).strip()
    password = str(data.get('password', ''))
    if len(password) < 6:
        return jsonify(error='新密码至少需要6位'), 400
    with db() as con:
        user = con.execute('SELECT id FROM users WHERE email=? AND phone=?', (email, phone)).fetchone()
        if not user:
            return jsonify(error='邮箱和手机号不匹配'), 400
        con.execute('UPDATE users SET password_hash=? WHERE id=?', (generate_password_hash(password), user['id']))
    return jsonify(ok=True)

@app.post('/api/auth/logout')
def logout():
    session.clear()
    return jsonify(ok=True)

@app.get('/api/transactions')
@login_required
def get_transactions():
    return jsonify(rows())

@app.post('/api/transactions')
@login_required
def add_transaction():
    data = request.get_json(force=True)
    required = ['kind','amount','category','tx_date']
    if any(k not in data for k in required): return jsonify(error='缺少必要字段'), 400
    category = data.get('category', '').strip()
    category = normalize_category(category, data.get('note', ''), data['kind'])
    with db() as con:
        cur = con.execute('INSERT INTO transactions(user_id,kind,amount,category,note,tx_date,created_at) VALUES(?,?,?,?,?,?,?)',
            (current_user_id(), data['kind'], float(data['amount']), category, data.get('note',''), data['tx_date'], datetime.now().isoformat()))
        return jsonify(id=cur.lastrowid)

@app.delete('/api/transactions/<int:tid>')
@login_required
def delete_transaction(tid):
    with db() as con: con.execute('DELETE FROM transactions WHERE id=? AND user_id=?', (tid, current_user_id()))
    return jsonify(ok=True)

@app.delete('/api/transactions')
@login_required
def clear_transactions():
    with db() as con:
        result = con.execute('DELETE FROM transactions WHERE user_id=?', (current_user_id(),))
    return jsonify(ok=True, deleted=result.rowcount)

IMPORT_ALIASES = {
    'kind': {'kind', 'type', '类型', '收支', '收支类型'},
    'amount': {'amount', '金额', '数额', '支出', '收入'},
    'category': {'category', '分类', '类别', '项目'},
    'note': {'note', '备注', '说明', '用途', '摘要'},
    'date': {'tx_date', 'date', '日期', '时间', '交易日期'}
}

def clean_cell(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value).strip() if value is not None else ''

def normalize_date(value):
    text = clean_cell(value)
    match = re.search(r'(20\d{2})(?:年(\d{1,2})月(\d{1,2})日?|[./-](\d{1,2})[./-](\d{1,2}))', text)
    if match:
        month, day = match.group(2) or match.group(4), match.group(3) or match.group(5)
        return f'{match.group(1)}-{int(month):02d}-{int(day):02d}'
    return ''

def amount_value(value):
    match = re.search(r'-?\d[\d,]*(?:\.\d+)?', clean_cell(value))
    return abs(float(match.group(0).replace(',', ''))) if match else None

def infer_kind(value, amount_text=''):
    text = f'{clean_cell(value)} {clean_cell(amount_text)}'.lower()
    return 'income' if any(word in text for word in ('收入', '工资', '薪资', '退款', '报销', 'income', '入账')) else 'expense'

def infer_category(text, kind='expense', fallback='未填写备注'):
    note_text = clean_cell(text)
    text = note_text.lower()
    if kind == 'income':
        for category, keywords in (
            ('项目收入', ('项目汇款', '项目结算', '项目收入', '服务项目', '客户结算', '项目款')),
            ('工资', ('工资', '薪资', '月薪', 'salary', 'payroll')),
            ('销售收入', ('销售', '订单到账', '线上订单', '收款')),
            ('兼职', ('兼职', '副业', 'freelance')),
            ('奖金', ('奖金', '年终奖', '奖励', 'bonus')),
            ('投资收益', ('投资收益', '理财收益', '基金收益', '股票收益', '分红')),
            ('退款', ('退款', '退货退款', 'refund')),
            ('经营收入', ('经营收入', '营业收入', '业务收入'))
        ):
            if any(keyword in text for keyword in keywords):
                return category
    keyword_groups = [
        ('餐饮', ('餐饮', '饮食', '吃饭', '午餐', '晚餐', '早餐', '外卖', '咖啡', '奶茶')),
        ('购物', ('购物', '消费', '淘宝', '京东', '拼多多', '超市', '衣服', '服装', '生活用品')),
        ('交通', ('交通', '地铁', '公交', '打车', '出租车', '通勤', '加油', '停车')),
        ('住房', ('住房', '房租', '租房', '物业', '水电', '燃气')),
        ('医疗', ('医疗', '看病', '医院', '药品', '买药')),
        ('教育', ('教育', '学费', '课程', '培训', '书籍')),
        ('娱乐', ('娱乐', '电影', '游戏', '旅游', '旅行', '门票')),
        ('通讯', ('通讯', '话费', '手机费', '宽带', '网费')),
        ('投资', ('投资', '理财', '基金', '股票', '分红')),
        ('采购', ('采购', '补货', '进货', '原材料')),
        ('办公', ('办公', '运营支出', '办公用品')),
        ('经营支出', ('经营支出', '日常经营', '固定支出'))
    ]
    for category, keywords in keyword_groups:
        if any(keyword in text for keyword in keywords):
            return category
    return note_text[:16] or fallback

def normalize_category(category, note='', kind='expense'):
    category = clean_cell(category)
    combined = f'{category} {clean_cell(note)}'.lower()
    if kind == 'expense' and any(word in combined for word in ('线上支付', '在线支付', '扫码支付', '银行卡消费', '刷卡消费', '支付消费')):
        return '支出'
    category_aliases = {
        '项目汇款': '项目收入',
        '项目款': '项目收入',
        '项目结算': '项目收入',
        '服务项目': '项目收入',
        '客户结算': '项目收入',
        '工资收入': '工资',
        '薪资': '工资',
        'salary': '工资',
        'payroll': '工资',
        '投资收入': '投资收益',
        '理财收益': '投资收益',
        '基金收益': '投资收益',
        '股票收益': '投资收益',
        '分红': '投资收益',
        '支出': '支出',
        '饮食': '餐饮',
        '午餐': '餐饮',
        '晚餐': '餐饮',
        '早餐': '餐饮',
        '外卖': '餐饮',
        '咖啡': '餐饮',
        '奶茶': '餐饮',
    }
    if category in category_aliases:
        category = category_aliases[category]
    elif category and category != category_aliases.get(category, category):
        category = category_aliases.get(category, category)
    income_categories = {'工资', '项目收入', '销售收入', '兼职', '奖金', '投资收益', '退款', '经营收入'}
    expense_categories = {'支出', '餐饮', '饮食', '购物', '交通', '住房', '医疗', '教育', '娱乐', '通讯', '投资', '采购', '办公', '经营支出'}
    if kind == 'income' and category not in income_categories:
        category = ''
    if kind == 'expense' and category not in expense_categories:
        category = ''
    if not category or category == '其他':
        return infer_category(note, kind)
    return category

def record_from_mapping(mapping):
    normalized = {clean_cell(key).lower(): clean_cell(value) for key, value in mapping.items()}
    def find(field):
        for alias in IMPORT_ALIASES[field]:
            if alias.lower() in normalized and normalized[alias.lower()]:
                return normalized[alias.lower()]
        return ''
    amount_text = find('amount')
    amount = amount_value(amount_text)
    if amount is None:
        return None
    tx_date = normalize_date(find('date'))
    if not tx_date:
        return None
    kind_text = find('kind')
    kind = 'income' if any(word in f'{kind_text} {amount_text}'.lower() for word in ('收入', '工资', '退款', 'income', '入账')) else 'expense'
    category = find('category')
    return {
        'kind': kind,
        'amount': amount,
        'category': normalize_category(category, f'{kind_text} {find("note")}', kind),
        'note': find('note'),
        'tx_date': tx_date
    }

def records_from_rows(rows):
    rows = [[clean_cell(cell) for cell in row] for row in rows if any(clean_cell(cell) for cell in row)]
    if not rows:
        return []
    header = [cell.lower() for cell in rows[0]]
    has_header = sum(any(alias.lower() == cell for alias in aliases) for cell in header for aliases in IMPORT_ALIASES.values()) >= 2
    if has_header:
        return [record for row in rows[1:] if (record := record_from_mapping(dict(zip(header, row))))]
    return records_from_text('\n'.join(' '.join(row) for row in rows))

def records_from_text(text):
    records = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        date_pattern = r'20\d{2}(?:年\d{1,2}月\d{1,2}日?|[./-]\d{1,2}[./-]\d{1,2})(?:[T\s]+\d{1,2}:\d{2}(?::\d{2})?)?'
        date_match = re.search(date_pattern, line)
        without_date = re.sub(date_pattern, '', line)
        amount = amount_value(without_date)
        if amount is None or not date_match:
            continue
        kind = infer_kind(line)
        date = normalize_date(date_match.group(0) if date_match else '')
        without_amount = re.sub(r'(?:¥|￥)?\s*-?\d[\d,]*(?:\.\d+)?', '', without_date)
        without_kind = re.sub(r'收入|支出|income|expense', '', without_amount, flags=re.I)
        parts = [part for part in re.split(r'[,，\t|\s]+', without_kind.strip()) if part]
        category = infer_category(line, kind, parts[0] if parts else '其他')
        records.append({'kind': kind, 'amount': amount, 'category': category, 'note': line, 'tx_date': date})
    return records

def parse_import_file(filename, raw):
    extension = os.path.splitext(filename.lower())[1]
    if extension == '.csv':
        return records_from_rows(list(csv.reader(io.StringIO(raw.decode('utf-8-sig', errors='ignore')))))
    if extension in ('.xlsx', '.xls'):
        if extension == '.xls':
            import xlrd
            workbook = xlrd.open_workbook(file_contents=raw)
            return records_from_rows(workbook.sheet_by_index(0).get_rows())
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        return records_from_rows(workbook.active.iter_rows(values_only=True))
    if extension == '.docx':
        document = Document(io.BytesIO(raw))
        rows = [[cell.text for cell in row.cells] for table in document.tables for row in table.rows]
        records = records_from_rows(rows)
        return records or records_from_text('\n'.join(paragraph.text for paragraph in document.paragraphs))
    if extension == '.pdf':
        reader = PdfReader(io.BytesIO(raw))
        return records_from_text('\n'.join(page.extract_text() or '' for page in reader.pages))
    raise ValueError('暂不支持该文件类型，请上传 CSV、Excel、Word 或 PDF 文件')

@app.post('/api/import')
@login_required
def import_file():
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify(error='未选择文件'), 400
    try:
        records = parse_import_file(f.filename, f.read())
    except Exception as error:
        return jsonify(error=f'文件解析失败：{error}'), 400
    if not records:
        return jsonify(error='没有识别到有效账目，请检查文档中是否包含日期、金额和分类信息'), 400
    with db() as con:
        con.executemany('INSERT INTO transactions(user_id,kind,amount,category,note,tx_date,created_at) VALUES(?,?,?,?,?,?,?)',
            [(current_user_id(), r['kind'], r['amount'], r['category'], r['note'], r['tx_date'], datetime.now().isoformat()) for r in records])
    return jsonify(imported=len(records))

def local_ai_answer(prompt, transactions):
    prompt = str(prompt or '').strip()
    income_items = [item for item in transactions if item['kind'] == 'income']
    expense_items = [item for item in transactions if item['kind'] == 'expense']
    income = sum(item['amount'] for item in income_items)
    expense = sum(item['amount'] for item in expense_items)
    balance = income - expense
    categories = {}
    for item in expense_items:
        categories[item['category']] = categories.get(item['category'], 0) + item['amount']
    ranked = sorted(categories.items(), key=lambda pair: pair[1], reverse=True)
    latest_month = max((str(item.get('tx_date', ''))[:7] for item in transactions), default='')
    recent = [item for item in transactions if str(item.get('tx_date', '')).startswith(latest_month)]
    recent_expense = sum(item['amount'] for item in recent if item['kind'] == 'expense')
    recent_income = sum(item['amount'] for item in recent if item['kind'] == 'income')
    money_text = lambda value: f'¥{value:,.2f}'
    if not transactions:
        return '目前账本还没有足够数据。我可以在你录入几笔收入和支出后，分析结余、分类占比、近期趋势和预算；现在最值得做的是先记录一周的真实开支。'
    top = ranked[0] if ranked else ('暂无支出分类', 0)
    top_ratio = top[1] / expense if expense else 0
    question = prompt.lower()
    if any(word in question for word in ('分类', '哪类', '最多', '支出结构', '花在哪')):
        details = '；'.join(f'{category} {money_text(amount)}（{amount / expense:.0%}）' for category, amount in ranked[:5]) if expense else '目前没有支出记录'
        return f'从现有 {len(expense_items)} 笔支出来看，支出主要集中在：{details}。其中 {top[0]} 是当前最需要关注的分类，占支出 {top_ratio:.0%}。建议先为它设置周限额，不要一开始平均压缩所有分类。'
    if any(word in question for word in ('收入', '结余', '剩下', '赚了')):
        return f'目前累计收入为 {money_text(income)}，累计支出为 {money_text(expense)}，结余为 {money_text(balance)}。最近有记录月份 {latest_month} 的收入为 {money_text(recent_income)}，支出为 {money_text(recent_expense)}。' + ('结余为正，可以先建立自动储蓄。' if balance > 0 else '当前支出已经超过收入，建议先暂停非必要支出并复核最近的大额分类。')
    if any(word in question for word in ('趋势', '最近', '本月', '这个月', '变化')):
        change = '最近月份支出高于收入' if recent_expense > recent_income else '最近月份收入高于支出'
        return f'按现有日期记录，最近有账目的月份是 {latest_month}：收入 {money_text(recent_income)}，支出 {money_text(recent_expense)}，{change}。累计来看，{top[0]}是最大支出来源；如果你想看更准确的上升或下降趋势，建议连续记录至少三个月。'
    if any(word in question for word in ('预算', '省钱', '怎么做', '建议', '规划', '控制')):
        available = max(balance, 0)
        saving = available * 0.5
        flexible = available * 0.3
        return f'结合当前账本，我建议下一个周期先把结余 {money_text(available)} 分成三部分：{money_text(saving)} 用于储蓄或应急金，{money_text(flexible)} 作为可变支出，剩余 {money_text(max(available - saving - flexible, 0))} 作为机动金。重点复盘 {top[0]}，每周达到该分类月均支出的 25% 时就暂停非必要消费。'
    return f'我看到你的账本累计有 {len(transactions)} 笔记录：收入 {money_text(income)}，支出 {money_text(expense)}，结余 {money_text(balance)}。当前最大支出分类是 {top[0]}（{money_text(top[1])}）。你可以继续问我“哪类支出最多”“最近趋势怎样”或“如何制定预算”，我会根据当前账本回答。'

@app.post('/api/ai')
@login_required
def ai():
    data = request.get_json(force=True)
    prompt = data.get('prompt','')
    provider = str(data.get('provider','demo')).lower()
    api_key = data.get('api_key','')
    provider, base_url, model = resolve_provider_settings(provider, data.get('base_url',''), data.get('model',''))
    tx = rows()
    income = sum(x['amount'] for x in tx if x['kind']=='income')
    expense = sum(x['amount'] for x in tx if x['kind']=='expense')
    summary = f'当前累计收入：{income:.2f}，累计支出：{expense:.2f}，结余：{income-expense:.2f}。'
    if provider in ('demo', 'local'):
        return jsonify(answer=local_ai_answer(prompt, tx), source='local-analysis')
    if provider == 'ollama':
        context = json.dumps({'账本摘要': summary, '交易记录': tx[-80:]}, ensure_ascii=False)
        payload = json.dumps({'model': model, 'messages': [
            {'role': 'system', 'content': '你是运行在用户本机的中文个人财务助手。像正常聊天一样理解用户问题，只根据账本数据回答，先给结论再给依据；不要套用固定欢迎词，不要编造数据，不提供保证收益的投资建议。'},
            {'role': 'user', 'content': context + '\n用户问题：' + prompt}
        ], 'temperature': 0.45}).encode()
        req = urllib.request.Request(base_url, data=payload, headers={'Content-Type': 'application/json'})
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                result = json.loads(resp.read().decode())
            return jsonify(answer=result['choices'][0]['message']['content'], source='ollama')
        except Exception as error:
            return jsonify(answer=local_ai_answer(prompt, tx), source='local-analysis', warning='本机 Ollama 暂不可用，已使用本地账本分析。', error=str(error)), 200
    if not api_key:
        return jsonify(answer=local_ai_answer(prompt, tx), source='local-analysis')
    payload = json.dumps({'model':model,'messages':[{'role':'system','content':'你是谨慎、客观的个人财务助手，不提供保证收益的投资建议。'}, {'role':'user','content':summary+'\n'+prompt}], 'temperature':0.3}).encode()
    req = urllib.request.Request(base_url, data=payload, headers={'Content-Type':'application/json','Authorization':'Bearer '+api_key})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode())
        return jsonify(answer=result['choices'][0]['message']['content'])
    except Exception as e:
        return jsonify(answer='AI 调用失败，请检查 API 地址、模型名称和密钥。', error=str(e)), 502

@app.post('/api/planning')
@login_required
def planning_api():
    data = request.get_json(force=True)
    provider = str(data.get('provider', 'demo')).lower()
    api_key = data.get('api_key', '')
    provider, base_url, model = resolve_provider_settings(provider, data.get('base_url', ''), data.get('model', ''))
    income = float(data.get('income', 0) or 0)
    save_rate = float(data.get('save_rate', 20) or 0)
    goal = str(data.get('goal', '')).strip() or '建立稳定的月度收支和储蓄计划'
    tx = rows()
    expenses = [item for item in tx if item['kind'] == 'expense']
    expense_by_category = {}
    for item in expenses:
        expense_by_category[item['category']] = expense_by_category.get(item['category'], 0) + item['amount']
    summary = {
        '收入': income,
        '已记录收入': sum(item['amount'] for item in tx if item['kind'] == 'income'),
        '已记录支出': sum(item['amount'] for item in expenses),
        '支出分类': expense_by_category,
        '储蓄目标比例': save_rate,
        '规划目标': goal
    }
    fallback = build_local_plan(summary, tx, income, save_rate, goal)
    if provider in ('demo', 'local'):
        return jsonify(answer=fallback, source='local-analysis', warning='未调用云端 API，已根据真实账目完成动态分析。')
    if provider == 'ollama':
        payload = json.dumps({'model': model, 'messages': [
            {'role': 'system', 'content': '你是运行在用户本机的个人财务规划助手。只根据输入数据分析，不编造账目；使用中文，给出详细、分阶段、可执行的规划，不提供保证收益的投资建议。'},
            {'role': 'user', 'content': '请根据以下真实账本摘要生成个性化规划，必须说明你观察到的收入稳定性、支出集中分类、近期变化、预算数字、分类行动和每周执行步骤：\n' + json.dumps(summary, ensure_ascii=False)}
        ], 'temperature': 0.35}).encode()
        req = urllib.request.Request(base_url, data=payload, headers={'Content-Type': 'application/json'})
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                result = json.loads(resp.read().decode())
            return jsonify(answer=result['choices'][0]['message']['content'], source='ollama')
        except Exception as error:
            return jsonify(answer=fallback, source='local-analysis', warning='本机 Ollama 当前不可用，已使用动态账本分析。请安装并运行 Ollama 后重试。', error=str(error)), 200
    if not api_key:
        return jsonify(answer=fallback, source='local-analysis', warning='当前服务商未配置 API Key，已使用动态账本分析。')
    prompt = '请根据以下真实账本摘要生成一份足够详尽、可执行的个人财务规划。必须包含：当前诊断、月度预算分配、分类限额建议、储蓄与应急金策略、每周和每月执行步骤、风险提醒。使用中文，金额保留两位小数，不要编造不存在的账目。\n' + json.dumps(summary, ensure_ascii=False)
    payload = json.dumps({'model': model, 'messages': [
        {'role': 'system', 'content': '你是谨慎、客观的个人财务规划助手，不提供保证收益的投资建议。'},
        {'role': 'user', 'content': prompt}
    ], 'temperature': 0.25}).encode()
    req = urllib.request.Request(base_url, data=payload, headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + api_key})
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            result = json.loads(resp.read().decode())
        return jsonify(answer=result['choices'][0]['message']['content'], source='deepseek')
    except Exception as error:
        return jsonify(answer=fallback, source='local', warning='DeepSeek 调用失败，已使用本地详细规划。', error=str(error)), 200

if __name__ == '__main__':
    init_db()
    app.run(debug=True, host='127.0.0.1', port=5000)
