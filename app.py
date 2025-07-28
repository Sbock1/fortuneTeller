import flask as fl
import sqlite3 as sql
import pandas as pd

app = fl.Flask(__name__)

def getDBconnection():
    conn = sql.connect("mainDatabase.db")
    conn.row_factory = sql.Row
    return conn

@app.route('/')
def index():
    with getDBconnection() as conn:
        stocks = conn.execute('SELECT * FROM Analysis_Quarterly LIMIT 10').fetchall()
    return fl.render_template('index.html', stocks=stocks)

@app.route('/filter', methods=['GET'])
def filter():
    symbol = fl.request.args.get('symbol')  # Symbol-Filter aus den Anfrageparametern
    
    # SQL-Abfrage vorbereiten
    query = """
        SELECT symbol, net_income, free_cashflow, earnings_per_share
        FROM Analysis_Quarterly
        WHERE symbol = ?
    """
    
    with getDBconnection() as conn:
        stocks = conn.execute(query, (symbol,)).fetchall()  # Filter nach Symbol
    
    return fl.render_template('index.html', stocks=stocks)