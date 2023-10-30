from flask import Flask, flash, redirect, render_template, request, session, jsonify, send_file
from cs50 import SQL
from flask_session import Session
from werkzeug.security import check_password_hash, generate_password_hash
from datetime import datetime
from helpers import login_required, change_excel_file


app = Flask(__name__)

# Configure session to use filesystem (instead of signed cookies)
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)


# Configure CS50 Library to use SQLite database
db = SQL("sqlite:///OrderTracking.db")

# Variables
AREA = {
        1: ['EMPU','IKEA TARGÓWEK - EMPU', 'ul. Malborska 53', '03-286 Warszawa'],
        2: ['SKLEP', 'IKEA TARGÓWEK - SKLEP', 'ul. Malborska 51', '03-286 Warszawa'],
        3: ['NoLimit', 'NO LIMIT', 'ul. Chełmżyńska 253A', '04-247 Warszawa']
    }

@app.after_request
def after_request(response):
    """Ensure responses aren't cached"""
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Expires"] = 0
    response.headers["Pragma"] = "no-cache"
    return response


@app.route("/", methods=["GET","POST"])
def index():

    if request.method == "POST":
        selected_date = request.form.get("selectedDate")
        
    else:
        now = datetime.now()
        current_date = now.date()
        selected_date = current_date.isoformat()

    query = """
    SELECT
        metastasis.*,
        users.username AS user_username
    FROM metastasis
    JOIN users ON metastasis.user_id = users.id
    WHERE metastasis.data = ?;
    """

    rows = db.execute(query, selected_date)

    return render_template("index.html", rows=rows, date=selected_date, area=AREA)



@app.route("/search", methods=["GET","POST"])
def search():

    if request.method == "POST":
        isell_number = request.form.get("search")
        
        #When user fails to provide number isell
        if isell_number == "":
            return render_template("search.html", error="Nie podano numeru zamówienia!")
        
        #Select all data from orders when user give number isell
        query = """
            SELECT
                users.id AS user_id,
                users.username AS user_username,
                metastasis.id AS metastasis_id,
                metastasis.data AS metastasis_data,
                metastasis.sent_from AS metastasis_sent_from,
                metastasis.sent_to AS metastasis_sent_to,
                orders.num_isell AS orders_num_isell
            FROM users
            JOIN metastasis ON users.id = metastasis.user_id
            JOIN orders ON metastasis.id = orders.metastasis_id
            WHERE orders.num_isell = ?;
            """

        rows = db.execute(query, isell_number)

        # if number isell does not exist return error
        if rows == []:
            return render_template("search.html", error=f"Brak historii dla zamówienia o numerze {isell_number}.")

        return render_template("search.html", rows=rows, area=AREA)

    return render_template("search.html", error="Nie podano numeru zamówienia!")


@app.route("/create", methods=["GET", "POST"])
@login_required
def create():
    if request.method == "POST":

        reg_num = request.form.get("reg_num")
        plomb_num = request.form.get("plomb_num")
        sent_to = int(request.form.get("sentTo"))
        sent_from = int(request.form.get("sentFrom"))
        isell_list = request.form.get("isellList")

        isell_list = isell_list.split("\n")
        isell_list = [i.strip() for i in isell_list]

        now = datetime.now()
        current_date = now.date()
        selected_date = current_date.isoformat()

        all_data = (
            session["user_id"],
            selected_date,
            f"{now.hour}:{now.minute}",
            reg_num,
            plomb_num,
            sent_from,
            sent_to
            )
        
        metastasis_id = db.execute(
            "INSERT INTO metastasis (user_id, data, time, reg_num, plomb_num, sent_from, sent_to) VALUES (?);", all_data)

        for isell in isell_list:
            db.execute(
                "INSERT INTO orders (metastasis_id, num_isell) VALUES (?);",
                (
                    metastasis_id,
                    isell
                )
            )

        user_username = db.execute(
            "SELECT username FROM users WHERE id = ?", session["user_id"])[0]["username"]


        change_excel_file(all_data, isell_list, AREA)

        return render_template("create.html", correct_message="Przerzut dodany!")

    return render_template("create.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    """Log user in"""

    # Forget any user_id
    session.clear()

    # User reached route via POST (as by submitting a form via POST)
    if request.method == "POST":
        # Ensure username was submitted
        if not request.form.get("username"):
            return render_template("login.html", error="Nie podano nazwy użytkownika!")

        # Ensure password was submitted
        elif not request.form.get("password"):
            return render_template("login.html", error="Nie podano hasła!")
        
        # Query database for username
        rows = db.execute(
            "SELECT * FROM users WHERE username = ?", request.form.get("username")
        )

        # Ensure username exists and password is correct
        if len(rows) != 1 or not check_password_hash(
            rows[0]["password"], request.form.get("password")
        ):
            return render_template("login.html", error="Niepoprawna nazwa użytkownika lub hasło.")

        # Remember which user has logged in
        session["user_id"] = rows[0]["id"]

        # Redirect user to home page
        return redirect("/")

    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    """Register user"""

    if request.method == "POST":
        if not request.form.get("username"):
            return render_template("register.html", error="Nie podano nazwy użytkownika!")

        if (
            db.execute(
                "SELECT * FROM users WHERE username = ?", request.form.get("username")
            )
            == []
        ):
            if not request.form.get("password"):
                return render_template("register.html", error="Nie podano hasła!")

            if not request.form.get("confirmation"):
                return render_template("register.html", error="Nie podano hasła!")

            if request.form.get("password") != request.form.get("confirmation"):
                return render_template("register.html", error="Hasła się różnią!")

            username = request.form.get("username")
            password = generate_password_hash(request.form.get("password"))
            db.execute(
                "INSERT INTO users (username, password) VALUES (?);", (username, password)
            )

            # Remember which user has logged in
            session["user_id"] = db.execute(
                "SELECT id FROM users WHERE username = ?", request.form.get("username")
            )[0]["id"]

            return redirect("/")

        return render_template("register.html", error="Użytkownik już istnieje!")

    return render_template("register.html")


@app.route('/download_excel', methods=['GET'])
def download_excel():
    excel_file = 'excel_file/list_przewozowy2.xlsx'
    return send_file(excel_file, as_attachment=True)