#!/usr/bin/env python3
"""
e-Διαχείριση - Οικονομική Διαχείριση
Διορθωμένη ημερομηνία (YYYY-MM-DD) για σωστό φιλτράρισμα & Excel
"""

import os
import sys
import webbrowser
import subprocess
import time
import socket
import threading
import traceback
from pathlib import Path
from datetime import datetime

# Ρύθμιση διαδρομών
APP_DIR = Path(__file__).parent.absolute()
DB_FILE = APP_DIR / "oikonomika.db"

print("=" * 60)
print("  💰 e-Διαχείριση - Εκκίνηση...")
print("=" * 60)

# Εγκατάσταση βιβλιοθηκών
print("\n📦 Έλεγχος βιβλιοθηκών...")
try:
    import flask
    import pandas
    import openpyxl

    print("   ✓ Όλες οι βιβλιοθήκες είναι εγκατεστημένες")
except ImportError as e:
    print(f"   Λείπει η βιβλιοθήκη: {e}")
    print("   📦 Εγκατάσταση...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'flask', 'pandas', 'openpyxl'])
    print("   ✓ Εγκαταστάθηκαν!")
    os.execv(sys.executable, [sys.executable] + sys.argv)

from flask import Flask, request, jsonify, send_file, render_template_string
import sqlite3
from io import BytesIO

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
                        <div class="mb-2">
                            <label>Ημερομηνία</label>
                            <input type="date" class="form-control" id="date" required>
                        </div>
                        <div class="mb-2">
                            <label>Περιγραφή</label>
                            <input type="text" class="form-control" id="desc" placeholder="π.χ. Σούπερ μάρκετ" required>
                        </div>
                        <div class="mb-2">
                            <label>Κατηγορία</label>
                            <select class="form-control" id="cat">
                                <option>Γενικά</option><option>Έσοδα</option><option>Διατροφή</option>
                                <option>Μεταφορές</option><option>Σπίτι</option><option>Ψυχαγωγία</option>
                            </select>
                        </div>
                        <div class="mb-2">
                            <label>Ποσό (+ για έσοδα, - για έξοδα)</label>
                            <input type="number" step="0.01" class="form-control" id="amount" placeholder="-50.00" required>
                        </div>
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
                            <input type="date" class="form-control form-control-sm" id="filterStart" placeholder="Από">
                            <span>έως</span>
                            <input type="date" class="form-control form-control-sm" id="filterEnd" placeholder="Έως">
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
        // Σήμερα σε YYYY-MM-DD για το input
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('date').value = today;
        // Προεπιλογή φίλτρου: αρχή μήνα έως σήμερα
        const firstDay = new Date(new Date().getFullYear(), new Date().getMonth(), 1)
            .toISOString().split('T')[0];
        document.getElementById('filterStart').value = firstDay;
        document.getElementById('filterEnd').value = today;

        function showMessage(msg, type) {
            const div = document.getElementById('message');
            div.innerHTML = `<div class="alert alert-${type}">${msg}</div>`;
            setTimeout(() => div.innerHTML = '', 3000);
        }

        // Δεν χρειάζεται μετατροπή – δουλεύουμε με YYYY-MM-DD παντού
        function getFilterDates() {
            return {
                start: document.getElementById('filterStart').value,
                end: document.getElementById('filterEnd').value
            };
        }

        // Εμφάνιση ημερομηνίας σε DD/MM/YYYY
        function formatDisplay(dateStr) {
            if (!dateStr) return '';
            const [y, m, d] = dateStr.split('-');
            return `${d}/${m}/${y}`;
        }

        async function loadData() {
            try {
                const { start, end } = getFilterDates();
                let url = '/api/data';
                if (start && end) {
                    url += `?start_date=${start}&end_date=${end}`;
                }

                const resp = await fetch(url);
                if (!resp.ok) throw new Error('HTTP ' + resp.status);
                const data = await resp.json();

                document.getElementById('balance').textContent = data.balance.toFixed(2) + '€';
                document.getElementById('income').textContent = data.income.toFixed(2) + '€';
                document.getElementById('expenses').textContent = Math.abs(data.expenses).toFixed(2) + '€';

                let html = '<table class="table table-striped"><thead><tr><th>Ημ/νία</th><th>Περιγραφή</th><th>Κατ.</th><th>Ποσό</th><th></th></tr></thead><tbody>';

                if (!data.transactions || data.transactions.length === 0) {
                    html += '<tr><td colspan="5" class="text-center text-muted">Καμία συναλλαγή</td></tr>';
                } else {
                    for (const t of data.transactions) {
                        const color = t.amount >= 0 ? 'amount-positive' : 'amount-negative';
                        const sign = t.amount >= 0 ? '+' : '';
                        // Η ημερομηνία έρχεται ως YYYY-MM-DD από τον server, την εμφανίζουμε DD/MM/YYYY
                        const displayDate = formatDisplay(t.date);
                        html += `<tr>
                            <td>${displayDate}</td>
                            <td>${t.description}</td>
                            <td><span class="badge bg-secondary">${t.category}</span></td>
                            <td class="${color}">${sign}${t.amount.toFixed(2)}€</td>
                            <td><button class="btn btn-sm btn-danger" onclick="deleteTransaction(${t.id})">🗑️</button></td>
                        </tr>`;
                    }
                }
                html += '</tbody></table>';
                document.getElementById('transactions').innerHTML = html;
            } catch (error) {
                document.getElementById('transactions').innerHTML = `<div class="alert alert-danger">Σφάλμα: ${error.message}</div>`;
            }
        }

        document.getElementById('form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const button = e.target.querySelector('button');
            button.disabled = true;
            button.textContent = 'Προσθήκη...';

            try {
                // Στέλνουμε την ημερομηνία ως YYYY-MM-DD (την τιμή του input)
                const data = {
                    date: document.getElementById('date').value,
                    description: document.getElementById('desc').value.trim(),
                    category: document.getElementById('cat').value,
                    amount: parseFloat(document.getElementById('amount').value)
                };

                const resp = await fetch('/api/add', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(data)
                });

                const result = await resp.json();
                if (!resp.ok) throw new Error(result.error || 'Σφάλμα στον server');

                showMessage('✅ Προστέθηκε!', 'success');
                document.getElementById('desc').value = '';
                document.getElementById('amount').value = '';
                await loadData();
            } catch (error) {
                showMessage('❌ ' + error.message, 'danger');
            } finally {
                button.disabled = false;
                button.textContent = 'Προσθήκη';
            }
        });

        async function deleteTransaction(id) {
            if (!confirm('Διαγραφή της συναλλαγής;')) return;
            await fetch('/api/del/' + id, {method: 'DELETE'});
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
            if (start && end) {
                url += `?start_date=${start}&end_date=${end}`;
            }
            window.location.href = url;
        }

        async function importExcel(input) {
            const file = input.files[0];
            if (!file) return;
            const formData = new FormData();
            formData.append('file', file);
            try {
                const resp = await fetch('/api/import', {method: 'POST', body: formData});
                const result = await resp.json();
                showMessage(result.message || 'Επιτυχία!', 'success');
                await loadData();
            } catch (error) {
                showMessage('❌ Σφάλμα εισαγωγής', 'danger');
            }
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
        conn.execute('''
                     CREATE TABLE IF NOT EXISTS transactions
                     (
                         id
                         INTEGER
                         PRIMARY
                         KEY
                         AUTOINCREMENT,
                         date
                         TEXT
                         NOT
                         NULL,
                         description
                         TEXT
                         NOT
                         NULL,
                         category
                         TEXT
                         DEFAULT
                         'Γενικά',
                         amount
                         REAL
                         NOT
                         NULL
                     )
                     ''')

        # Έλεγχος για στήλη category (παλαιότερες εκδόσεις)
        cursor = conn.execute("PRAGMA table_info(transactions)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'category' not in columns:
            print("⚙️  Προσθήκη στήλης category...")
            conn.execute("ALTER TABLE transactions ADD COLUMN category TEXT DEFAULT 'Γενικά'")
            print("✓ Η στήλη category προστέθηκε!")

        # Μετατροπή υπαρχουσών ημερομηνιών από DD/MM/YYYY σε YYYY-MM-DD (αν υπάρχουν)
        # Ελέγχουμε αν κάποια εγγραφή έχει "/" (παλιά μορφή)
        needs_migration = False
        try:
            sample = conn.execute("SELECT date FROM transactions LIMIT 1").fetchone()
            if sample and '/' in sample[0]:
                needs_migration = True
        except:
            pass

        if needs_migration:
            print("🔄 Μετατροπή ημερομηνιών σε YYYY-MM-DD...")
            rows = conn.execute("SELECT id, date FROM transactions").fetchall()
            for row in rows:
                old_date = row[1]
                if '/' in old_date:
                    parts = old_date.split('/')
                    if len(parts) == 3:
                        new_date = f"{parts[2]}-{parts[1]}-{parts[0]}"
                        conn.execute("UPDATE transactions SET date = ? WHERE id = ?", (new_date, row[0]))
            conn.commit()
            print("✓ Οι ημερομηνίες μετατράπηκαν επιτυχώς.")

        conn.commit()
        conn.close()
        print("✓ Βάση δεδομένων έτοιμη")
        return True
    except Exception as e:
        print(f"❌ Σφάλμα βάσης δεδομένων: {e}")
        traceback.print_exc()
        return False


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/api/data')
def get_data():
    try:
        conn = sqlite3.connect(str(DB_FILE))
        conn.row_factory = sqlite3.Row

        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')

        query = "SELECT * FROM transactions"
        params = []
        if start_date and end_date:
            query += " WHERE date BETWEEN ? AND ?"
            params = [start_date, end_date]
        query += " ORDER BY date DESC"

        transactions = [dict(row) for row in conn.execute(query, params).fetchall()]

        sum_query = "SELECT COALESCE(SUM(amount), 0) FROM transactions"
        sum_params = []
        if start_date and end_date:
            sum_query += " WHERE date BETWEEN ? AND ?"
            sum_params = [start_date, end_date]

        total_balance = conn.execute(sum_query, sum_params).fetchone()[0]

        income_query = "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE amount > 0"
        expense_query = "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE amount < 0"
        if start_date and end_date:
            income_query += " AND date BETWEEN ? AND ?"
            expense_query += " AND date BETWEEN ? AND ?"
            total_income = conn.execute(income_query, [start_date, end_date]).fetchone()[0]
            total_expenses = conn.execute(expense_query, [start_date, end_date]).fetchone()[0]
        else:
            total_income = conn.execute(income_query).fetchone()[0]
            total_expenses = conn.execute(expense_query).fetchone()[0]

        conn.close()

        return jsonify({
            'transactions': transactions,
            'balance': float(total_balance),
            'income': float(total_income),
            'expenses': float(total_expenses)
        })
    except Exception as e:
        print(f"❌ Σφάλμα στο /api/data: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/add', methods=['POST'])
def add_transaction():
    try:
        data = request.get_json(force=True)
        print(f"📥 Λήφθηκαν δεδομένα: {data}")

        if not data:
            return jsonify({'error': 'Δεν εστάλησαν δεδομένα'}), 400

        date = data.get('date', '')  # YYYY-MM-DD
        description = data.get('description', '')
        category = data.get('category', 'Γενικά')

        try:
            amount = float(data.get('amount', 0))
        except (ValueError, TypeError):
            return jsonify({'error': f'Μη έγκυρο ποσό: {data.get("amount")}'}), 400

        if not date or not description:
            return jsonify({'error': 'Η ημερομηνία και η περιγραφή είναι υποχρεωτικά'}), 400

        # Προαιρετική επικύρωση μορφής YYYY-MM-DD
        if len(date) != 10 or date[4] != '-' or date[7] != '-':
            return jsonify({'error': 'Η ημερομηνία πρέπει να είναι στη μορφή YYYY-MM-DD'}), 400

        conn = sqlite3.connect(str(DB_FILE))
        cursor = conn.execute(
            "INSERT INTO transactions (date, description, category, amount) VALUES (?, ?, ?, ?)",
            (date, description, category, amount)
        )
        conn.commit()
        new_id = cursor.lastrowid
        conn.close()

        print(f"✅ Προστέθηκε! ID: {new_id}")
        return jsonify({'success': True, 'id': new_id, 'message': 'Η συναλλαγή προστέθηκε!'}), 201

    except Exception as e:
        print(f"❌ ΣΦΑΛΜΑ στην προσθήκη: {e}")
        traceback.print_exc()
        return jsonify({'error': f'Σφάλμα server: {str(e)}'}), 500


@app.route('/api/del/<int:id>', methods=['DELETE'])
def delete_transaction(id):
    try:
        conn = sqlite3.connect(str(DB_FILE))
        conn.execute("DELETE FROM transactions WHERE id = ?", (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/export')
def export_excel():
    try:
        import pandas as pd
        conn = sqlite3.connect(str(DB_FILE))

        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')

        query = "SELECT date as 'Ημερομηνία', description as 'Περιγραφή', category as 'Κατηγορία', amount as 'Ποσό' FROM transactions"
        params = []
        if start_date and end_date:
            query += " WHERE date BETWEEN ? AND ?"
            params = [start_date, end_date]
        query += " ORDER BY date DESC"

        df = pd.read_sql_query(query, conn, params=params)
        conn.close()

        # Μετατροπή της στήλης ημερομηνίας σε DD/MM/YYYY για το Excel
        def convert_date(d):
            if isinstance(d, str) and '-' in d:
                y, m, day = d.split('-')
                return f"{day}/{m}/{y}"
            return d

        if 'Ημερομηνία' in df.columns:
            df['Ημερομηνία'] = df['Ημερομηνία'].apply(convert_date)

        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'oikonomika_{datetime.now().strftime("%Y%m%d")}.xlsx'
        )
    except Exception as e:
        print(f"❌ Σφάλμα εξαγωγής: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/import', methods=['POST'])
def import_excel():
    try:
        import pandas as pd

        if 'file' not in request.files:
            return jsonify({'error': 'Δεν επιλέχθηκε αρχείο'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Κενό αρχείο'}), 400

        df = pd.read_excel(file, engine='openpyxl')

        conn = sqlite3.connect(str(DB_FILE))
        imported = 0
        for _, row in df.iterrows():
            try:
                # Αν η ημερομηνία είναι ήδη string, πιθανόν DD/MM/YYYY -> μετατροπή σε YYYY-MM-DD
                date_val = str(row['Ημερομηνία'])
                if '/' in date_val:
                    parts = date_val.split('/')
                    if len(parts) == 3:
                        date_val = f"{parts[2]}-{parts[1]}-{parts[0]}"
                conn.execute(
                    "INSERT INTO transactions (date, description, category, amount) VALUES (?, ?, ?, ?)",
                    (
                        date_val,
                        str(row['Περιγραφή']),
                        str(row.get('Κατηγορία', 'Γενικά')),
                        float(row['Ποσό'])
                    )
                )
                imported += 1
            except:
                continue

        conn.commit()
        conn.close()

        return jsonify({'success': True, 'message': f'Εισήχθησαν {imported} εγγραφές!'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("\n🔧 Αρχικοποίηση...")

    if not init_db():
        print("❌ Σφάλμα στη βάση δεδομένων")
        input("Πάτησε Enter για έξοδο...")
        sys.exit(1)

    port = 5000
    url = f'http://localhost:{port}'

    print(f"\n{'=' * 60}")
    print(f"  🌐 Η εφαρμογή είναι έτοιμη!")
    print(f"  📍 Διεύθυνση: {url}")
    print(f"  💡 Δες αυτό το παράθυρο για σφάλματα")
    print(f"{'=' * 60}\n")


    def open_browser():
        time.sleep(1.5)
        webbrowser.open(url)


    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='127.0.0.1', port=port, debug=True)