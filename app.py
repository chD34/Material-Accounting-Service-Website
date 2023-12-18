from flask import Flask, render_template, redirect, url_for, request, flash, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from flask import send_file
from openpyxl import Workbook
from io import BytesIO


app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///site.db'
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(20), unique=True, nullable=False)
    name = db.Column(db.String(20), unique=True, nullable=False)
    surname = db.Column(db.String(20), unique=True, nullable=False)
    password = db.Column(db.String(60), nullable=False)
    position = db.Column(db.String(50), nullable=False)


class MaterialOperation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    username = db.Column(db.String(50), nullable=False)
    position = db.Column(db.String(50), nullable=False)
    subject = db.Column(db.String(50), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    sender = db.Column(db.String(50), nullable=False)
    receiver = db.Column(db.String(50), nullable=False)
    action = db.Column(db.String(20), nullable=False)

    timestamp = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password):
            login_user(user)
            jsonify({'message' : 'Login successful!', 'status' : 'Success'})
            return redirect(url_for('index'))
        else:
            jsonify({'message' : 'Login failed. Check your username and password.', 'status' : 'error'})
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('index'))


@app.route('/register', methods=['GET', 'POST'])
def register():
    positions = ['Адміністратор', 'Розробник', 'Тестувальник', 'Бізнес-аналітик', 'Фінансовий-аналітик']
    if request.method == 'POST':
        username = request.form.get('username')
        name = request.form.get('name')
        surname = request.form.get('surname')
        password = request.form.get('password')
        position = request.form.get('position')

        existing_user = User.query.filter_by(username=username).first()
        if existing_user:
            flash('Цей логін вже зарєстрований. Будь-ласка оберіть інший.', 'danger')
        else:
            hashed_password = generate_password_hash(password, method='pbkdf2:sha256', salt_length=8)
            new_user = User(username=username, name = name, surname = surname, password=hashed_password, position=position)
            db.session.add(new_user)
            db.session.commit()
            flash('Успішна реєстрація!', 'success')
            return redirect(url_for('login'))

    return render_template('register.html', positions=positions)


@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', user=current_user)


@app.route('/profile')
@login_required
def profile():
    return render_template('profile.html', user=current_user)


@app.route('/update_profile', methods=['POST'])
@login_required
def update_profile():
    current_user.name = request.form['name']
    current_user.surname = request.form['surname']
    db.session.commit()

    return redirect(url_for('profile'))


@app.route('/material_accounting', methods=['GET', 'POST'])
@login_required
def material_accounting():
    if request.method == 'POST':
        subject = request.form.get('subject')
        quantity = int(request.form.get('quantity'))
        sender = request.form.get('sender')
        receiver = request.form.get('receiver')
        action = request.form.get('action')

        material_operation = MaterialOperation(
            user_id=current_user.id,
            username = current_user.username,
            position = current_user.position,
            subject=subject,
            quantity=quantity,
            sender=sender,
            receiver=receiver,
            action=action
        )

        db.session.add(material_operation)
        db.session.commit()
        flash(f'Material operation ({action}) recorded successfully!', 'success')

    operations = MaterialOperation.query.all()
    return render_template('material_accounting.html', operations=operations)


@app.route('/export_data')
@login_required
def export_data():
    data = [['Користувач', 'Посада', 'Предмет', 'Кількість', 'Постачальник', 'Отримувач', 'Дія', 'Timestamp']]
    operations = MaterialOperation.query.all()

    for operation in operations:
        data.append([operation.username, operation.position, operation.subject, operation.quantity, operation.sender, operation.receiver, operation.action, operation.timestamp])

    workbook = Workbook()
    sheet = workbook.active

    for row in data:
        sheet.append(row)

    excel_data = BytesIO()
    workbook.save(excel_data)
    excel_data.seek(0)

    return send_file(excel_data, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='material_accounting_operations.xlsx')


@app.route('/all_users')
@login_required
def all_users():
    users = User.query.all()
    return render_template('all_users.html', users=users)


if __name__ == '__main__':
    with app.app_context():
        # db.drop_all()
        db.create_all()
    app.run(debug=True)

