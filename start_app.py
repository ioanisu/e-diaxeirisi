#!/usr/bin/env python3
"""
e-Διαχείριση - Αυτόνομη εφαρμογή
Δεν απαιτεί εγκατάσταση Python από τον χρήστη.
"""

import os
import sys
import webbrowser
import time
import socket
import threading
import traceback
from pathlib import Path
from datetime import datetime

# Εντοπισμός φακέλου (δουλεύει και μέσα στο PyInstaller bundle)
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys._MEIPASS)
else:
    APP_DIR = Path(__file__).parent.absolute()

DB_FILE = Path.home() / "e-diaxeirisi.db"  # αποθήκευση στον φάκελο του χρήστη

# Εισαγωγή βιβλιοθηκών (θα υπάρχουν ήδη στο bundle)
from flask import Flask, request, jsonify, send_file, render_template_string
import sqlite3
from io import BytesIO
import pandas as pd
import openpyxl

app = Flask(__name__)

HTML = '''
<!DOCTYPE html>
<html lang="el">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>💰 e-Διαχείριση</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f0f2f5; font-family: 'Segoe UI', sans-serif; }
        .navbar { background: linear-gradient(135deg, #667eea, #764ba2); color: white; }
        .card { background: white; border-radius: 15px; padding: 20px; margin: 10px 0; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .btn-primary { background: #667eea; border: none; }
        .btn-primary:hover { background: #5a6fd6; }
        .amount-positive { color: #10b981; font-weight: bold; }
        .amount-negative { color: #ef4444; font-weight: bold; }
        .filter-box { background: #f8f9fa; border-radius: 10px; padding: 10px; margin-bottom: 10px; }
    </style>
</head>
<body>
    <nav class="navbar p-3">
        <div class="container">
            <h3>💰 e-Διαχείριση</h3>
            <div>
                <button class="btn btn-light btn-sm me-2" onclick="exportExcel()">📥 Excel</button>
                <button class="btn btn-light btn-sm" onclick="document.getElementById('fileInput').click()">📤 Εισαγωγή</button>
                <input type="file" id="fileInput" accept=".xlsx" style="display:none" onchange="importExcel(this)">
            </div>
        </div>
    </nav>

    <div class="container mt-4">
        <div id="message"></div>
        <div class="row">
            <div class="col-md-4">
                <div class="card">
                    <h5>➕ Νέα Συναλλαγή</h5>
                    <form id="form">
                        <div class="mb-2"><label>Ημερομηνία</label><input type="date" class="form-control" id="date" required></div>
                        <div class="mb-2"><label>Περιγραφή</label><input type="text" class="form-control" id="desc" required></div>
                        <div class="mb-2"><label>Κατηγορία</label><select class="form-control" id="cat"><option>Γενικά</option><option>Έσοδα</option><option>Διατροφή</option><option>Μεταφορές</option><option>Σπίτι</option><option>Ψυχαγωγία</option></select></div>
                        <div class="mb-2"><label>Ποσό (+ έσοδο, - έξοδο)</label><input type="number" step="0.01" class="form-control" id="amount" required></div>
                        <button type="submit" class="btn btn-primary w-100">Προσθήκη</button>
                    </form>
                </div>
                <div class="card">
                    <h5>📊 Σύνοψη</h5>
                    <p>Υπόλοιπο: <strong id="balance">0.00€</strong></p>
                    <p>Έσοδα: <strong class="amount-positive" id="income">0.00€</strong></p>
                    <p>Έξοδα: <strong class="amount-negative" id="expenses">0.00€</strong></p>
                </div>
            </div>
            <div class="col-md-8">
                <div class="card">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h5>📋 Συναλλαγές</h5>
                        <div class="filter-box d-flex align-items-center gap-2">
                            <input type="date" class="form-control form-control-sm" id="filterStart">
                            <span>έως</span>
                            <input type="date" class="form-control form-control-sm" id="filterEnd">
                            <button class="btn btn-sm btn-outline-primary" onclick="applyFilter()">Φίλτρο</button>
                            <button class="btn btn-sm btn-outline-secondary" onclick="clearFilter()">Καθαρισμός</button>
                        </div>
                    </div>
                    <div id="transactions">Φόρτωση...</div>
                </div>
            </div>
        </div>
    </div>
    <script>
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('date').value = today;
        document.getElementById('filterStart').value = new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().split('T')[0];
        document.getElementById('filterEnd').value = today;

        function showMessage(msg, type) {
            const div = document.getElementById('message');
            div.innerHTML = `<div class="alert alert-${type}">${msg}</div>`;
            setTimeout(() => div.innerHTML = '', 3000);
        }

        function formatDisplay(d) { if(!d) return ''; const [y,m,day]=d.split('-'); return `${day}/${m}/${y}`; }
        function getFilterDates() { return { start: document.getElementById('filterStart').value, end: document.getElementById('filterEnd').value }; }

        async function loadData() {
            try {
                const { start, end } = getFilterDates();
                let url = '/api/data';
                if (start && end) url += `?start_date=${start}&end_date=${end}`;
                const resp = await fetch(url);
                if (!resp.ok) throw new Error('HTTP ' + resp.status);
                const data = await resp.json();
                document.getElementById('balance').textContent = data.balance.toFixed(2)+'€';
                document.getElementById('income').textContent = data.income.toFixed(2)+'€';
                document.getElementById('expenses').textContent = Math.abs(data.expenses).toFixed(2)+'€';
                let html = '<table class="table table-striped"><thead><tr><th>Ημ/νία</th><th>Περιγραφή</th><th>Κατ.</th><th>Ποσό</th><th></th></tr></thead><tbody>';
                if (!data.transactions || data.transactions.length===0) html += '<tr><td colspan="5" class="text-center text-muted">Καμία συναλλαγή</td></tr>';
                else for (const t of data.transactions) {
                    const color = t.amount>=0 ? 'amount-positive' : 'amount-negative';
                    const sign = t.amount>=0 ? '+' : '';
                    html += `<tr><td>${formatDisplay(t.date)}</td><td>${t.description}</td><td><span class="badge bg-secondary">${t.category}</span></td><td class="${color}">${sign}${t.amount.toFixed(2)}€</td><td><button class="btn btn-sm btn-danger" onclick="deleteTransaction(${t.id})">🗑️</button></td></tr>`;
                }
                html += '</tbody></table>';
                document.getElementById('transactions').innerHTML = html;
            } catch (error) {
                document.getElementById('transactions').innerHTML = `<div class="alert alert-danger">Σφάλμα: ${error.message}</div>`;
            }
        }

        document.getElementById('form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const btn = e.target.querySelector('button');
            btn.disabled = true; btn.textContent = 'Προσθήκη...';
            try {
                const data = {
                    date: document.getElementById('date').value,
                    description: document.getElementById('desc').value.trim(),
                    category: document.getElementById('cat').value,
                    amount: parseFloat(document.getElementById('amount').value)
                };
                const resp = await fetch('/api/add', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)});
                const result = await resp.json();
                if (!resp.ok) throw new Error(result.error || 'Σφάλμα');
                showMessage('✅ Προστέθηκε!', 'success');
                document.getElementById('desc').value = '';
                document.getElementById('amount').value = '';
                await loadData();
            } catch (error) {
                showMessage('❌ ' + error.message, 'danger');
            } finally { btn.disabled = false; btn.textContent = 'Προσθήκη'; }
        });

        async function deleteTransaction(id) {
            if (!confirm('Διαγραφή;')) return;
            await fetch('/api/del/'+id, {method:'DELETE'});
            showMessage('🗑️ Διαγράφηκε!', 'success');
            await loadData();
        }
        function applyFilter() { loadData(); }
        function clearFilter() {
            document.getElementById('filterStart').value = '';
            document.getElementById('filterEnd').value = '';
            loadData();
        }
        function exportExcel() {
            const { start, end } = getFilterDates();
            let url = '/api/export';
            if (start && end) url += `?start_date=${start}&end_date=${end}`;
            window.location.href = url;
        }
        async function importExcel(input) {
            const file = input.files[0];
            if (!file) return;
            const formData = new FormData(); formData.append('file', file);
            try {
                const resp = await fetch('/api/import', {method:'POST', body:formData});
                const result = await resp.json();
                showMessage(result.message || 'Επιτυχία!', 'success');
                await loadData();
            } catch (error) { showMessage('❌ Σφάλμα εισαγωγής', 'danger'); }
            input.value = '';
        }
        loadData();
    </script>
</body>
</html>
'''

def init_db():
    try:
        conn = sqlite3.connect(str(DB_FILE))
        conn.execute('''CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            description TEXT NOT NULL,
            category TEXT DEFAULT 'Γενικά',
            amount REAL NOT NULL)''')
        cursor = conn.execute("PRAGMA table_info(transactions)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'category' not in columns:
            conn.execute("ALTER TABLE transactions ADD COLUMN category TEXT DEFAULT 'Γενικά'")
        # Μετανάστευση ημερομηνιών αν υπάρχουν σε DD/MM/YYYY
        rows = conn.execute("SELECT id, date FROM transactions").fetchall()
        for row in rows:
            if '/' in row[1]:
                parts = row[1].split('/')
                if len(parts)==3:
                    conn.execute("UPDATE transactions SET date = ? WHERE id = ?", (f"{parts[2]}-{parts[1]}-{parts[0]}", row[0]))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/api/data')
def get_data():
    conn = sqlite3.connect(str(DB_FILE))
    conn.row_factory = sqlite3.Row
    start = request.args.get('start_date')
    end = request.args.get('end_date')
    query = "SELECT * FROM transactions"
    params = []
    if start and end:
        query += " WHERE date BETWEEN ? AND ?"
        params = [start, end]
    query += " ORDER BY date DESC"
    transactions = [dict(row) for row in conn.execute(query, params).fetchall()]
    sum_q = "SELECT COALESCE(SUM(amount),0) FROM transactions"
    if start and end:
        sum_q += " WHERE date BETWEEN ? AND ?"
        balance = conn.execute(sum_q, [start, end]).fetchone()[0]
        income = conn.execute("SELECT COALESCE(SUM(amount),0) FROM transactions WHERE amount>0 AND date BETWEEN ? AND ?", [start, end]).fetchone()[0]
        expenses = conn.execute("SELECT COALESCE(SUM(amount),0) FROM transactions WHERE amount<0 AND date BETWEEN ? AND ?", [start, end]).fetchone()[0]
    else:
        balance = conn.execute(sum_q).fetchone()[0]
        income = conn.execute("SELECT COALESCE(SUM(amount),0) FROM transactions WHERE amount>0").fetchone()[0]
        expenses = conn.execute("SELECT COALESCE(SUM(amount),0) FROM transactions WHERE amount<0").fetchone()[0]
    conn.close()
    return jsonify({'transactions':transactions, 'balance':float(balance), 'income':float(income), 'expenses':float(expenses)})

@app.route('/api/add', methods=['POST'])
def add():
    data = request.get_json(force=True)
    date = data.get('date')
    desc = data.get('description')
    cat = data.get('category', 'Γενικά')
    try:
        amount = float(data.get('amount', 0))
    except:
        return jsonify({'error':'Μη έγκυρο ποσό'}), 400
    if not date or not desc:
        return jsonify({'error':'Συμπληρώστε όλα τα πεδία'}), 400
    conn = sqlite3.connect(str(DB_FILE))
    cur = conn.execute("INSERT INTO transactions (date, description, category, amount) VALUES (?,?,?,?)", (date, desc, cat, amount))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return jsonify({'success':True, 'id':new_id}), 201

@app.route('/api/del/<int:id>', methods=['DELETE'])
def delete(id):
    conn = sqlite3.connect(str(DB_FILE))
    conn.execute("DELETE FROM transactions WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return jsonify({'success':True})

@app.route('/api/export')
def export():
    conn = sqlite3.connect(str(DB_FILE))
    start = request.args.get('start_date')
    end = request.args.get('end_date')
    query = "SELECT date as 'Ημερομηνία', description as 'Περιγραφή', category as 'Κατηγορία', amount as 'Ποσό' FROM transactions"
    params = []
    if start and end:
        query += " WHERE date BETWEEN ? AND ?"
        params = [start, end]
    query += " ORDER BY date DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    if 'Ημερομηνία' in df.columns:
        df['Ημερομηνία'] = df['Ημερομηνία'].apply(lambda d: f"{d[8:10]}/{d[5:7]}/{d[:4]}" if isinstance(d, str) and len(d)==10 and d[4]=='-' else d)
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'oikonomika_{datetime.now().strftime("%Y%m%d")}.xlsx')

@app.route('/api/import', methods=['POST'])
def import_excel():
    file = request.files['file']
    df = pd.read_excel(file, engine='openpyxl')
    conn = sqlite3.connect(str(DB_FILE))
    imported = 0
    for _, row in df.iterrows():
        try:
            date_val = str(row['Ημερομηνία'])
            if '/' in date_val:
                parts = date_val.split('/')
                if len(parts)==3:
                    date_val = f"{parts[2]}-{parts[1]}-{parts[0]}"
            conn.execute("INSERT INTO transactions (date, description, category, amount) VALUES (?,?,?,?)",
                         (date_val, str(row['Περιγραφή']), str(row.get('Κατηγορία','Γενικά')), float(row['Ποσό'])))
            imported += 1
        except:
            continue
    conn.commit()
    conn.close()
    return jsonify({'success':True, 'message':f'Εισήχθησαν {imported} εγγραφές'})

if __name__ == '__main__':
    if not init_db():
        print("Σφάλμα βάσης")
        sys.exit(1)
    port = 5000
    url = f'http://127.0.0.1:{port}'
    threading.Thread(target=lambda: webbrowser.open(url), daemon=True).start()
    app.run(host='127.0.0.1', port=port, debug=False)
