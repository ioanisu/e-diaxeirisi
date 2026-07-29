#!/usr/bin/env python3
"""
e-Διαχείριση - Αυτόνομη Εφαρμογή
Δεν χρειάζεται εγκατάσταση. Απλά διπλό κλικ!
Δημιουργεί αυτόματα φάκελο στην επιφάνεια εργασίας.
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

# ─── ΑΥΤΟΜΑΤΟΣ ΕΝΤΟΠΙΣΜΟΣ ΕΠΙΦΑΝΕΙΑΣ ΕΡΓΑΣΙΑΣ ───
def get_desktop_path():
    """Βρίσκει την επιφάνεια εργασίας σε Windows/macOS/Linux"""
    if sys.platform == 'win32':
        import ctypes.wintypes
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, 0, None, 0, buf)
        return Path(buf.value)
    elif sys.platform == 'darwin':  # macOS
        return Path.home() / 'Desktop'
    else:  # Linux
        return Path.home() / 'Desktop'

# Δημιουργία φακέλου δεδομένων στην επιφάνεια εργασίας
DESKTOP = get_desktop_path()
DATA_FOLDER = DESKTOP / 'e-Diaxeirisi_Data'
DATA_FOLDER.mkdir(exist_ok=True)

# Αρχείο βάσης δεδομένων (ΠΟΤΕ δεν χάνεται!)
DB_FILE = DATA_FOLDER / 'oikonomika.db'

# Φάκελος για εξαγόμενα Excel
EXPORTS_FOLDER = DATA_FOLDER / 'Exports'
EXPORTS_FOLDER.mkdir(exist_ok=True)

print("=" * 60)
print("  💰 e-Διαχείριση - Αυτόνομη Εφαρμογή")
print("=" * 60)
print(f"📂 Τα δεδομένα αποθηκεύονται στο:")
print(f"   {DATA_FOLDER}")
print(f"📁 Τα Excel εξάγονται στο:")
print(f"   {EXPORTS_FOLDER}")
print("=" * 60)

# Εισαγωγή βιβλιοθηκών
try:
    from flask import Flask, request, jsonify, send_file, render_template_string
    import sqlite3
    from io import BytesIO
    import pandas as pd
    import openpyxl
except ImportError as e:
    print(f"❌ Λείπει βιβλιοθήκη: {e}")
    print("📦 Προσπάθεια αυτόματης εγκατάστασης...")
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'flask', 'pandas', 'openpyxl'])
    print("✓ Οι βιβλιοθήκες εγκαταστάθηκαν. Παρακαλώ επανεκκινήστε την εφαρμογή.")
    input("Πατήστε Enter για έξοδο...")
    sys.exit(0)

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
        .info-banner { background: #e8f4fd; border-left: 4px solid #667eea; padding: 10px; margin-bottom: 20px; border-radius: 5px; }
    </style>
</head>
<body>
    <nav class="navbar p-3">
        <div class="container">
            <h3>💰 e-Διαχείριση</h3>
            <div>
                <button class="btn btn-light btn-sm me-2" onclick="exportExcel()">📥 Εξαγωγή Excel</button>
                <button class="btn btn-light btn-sm" onclick="document.getElementById('fileInput').click()">📤 Εισαγωγή Excel</button>
                <input type="file" id="fileInput" accept=".xlsx" style="display:none" onchange="importExcel(this)">
            </div>
        </div>
    </nav>

    <div class="container mt-4">
        <div id="message"></div>
        
        <div class="info-banner">
            <strong>💡 Τα δεδομένα σας αποθηκεύονται αυτόματα!</strong> 
            Μπορείτε να κλείσετε τον browser οποιαδήποτε στιγμή χωρίς να χαθεί τίποτα.
        </div>
        
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
                                <option>Υγεία</option><option>Εκπαίδευση</option>
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
                    <hr>
                    <small class="text-muted">📂 Τα δεδομένα αποθηκεύονται στην επιφάνεια εργασίας σας</small>
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

        function formatDisplay(d) { 
            if(!d) return ''; 
            const parts = d.split('-');
            return parts[2] + '/' + parts[1] + '/' + parts[0];
        }
        
        function getFilterDates() { 
            return { 
                start: document.getElementById('filterStart').value, 
                end: document.getElementById('filterEnd').value 
            }; 
        }

        async function loadData() {
            try {
                const { start, end } = getFilterDates();
                let url = '/api/data';
                if (start && end) url += `?start_date=${start}&end_date=${end}`;
                
                const resp = await fetch(url);
                if (!resp.ok) throw new Error('HTTP ' + resp.status);
                const data = await resp.json();
                
                document.getElementById('balance').textContent = data.balance.toFixed(2) + '€';
                document.getElementById('income').textContent = data.income.toFixed(2) + '€';
                document.getElementById('expenses').textContent = Math.abs(data.expenses).toFixed(2) + '€';
                
                let html = '<table class="table table-striped"><thead><tr><th>Ημ/νία</th><th>Περιγραφή</th><th>Κατ.</th><th>Ποσό</th><th></th></tr></thead><tbody>';
                
                if (!data.transactions || data.transactions.length === 0) {
                    html += '<tr><td colspan="5" class="text-center text-muted">Καμία συναλλαγή ακόμα</td></tr>';
                } else {
                    for (const t of data.transactions) {
                        const color = t.amount >= 0 ? 'amount-positive' : 'amount-negative';
                        const sign = t.amount >= 0 ? '+' : '';
                        html += `<tr>
                            <td>${formatDisplay(t.date)}</td>
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
                document.getElementById('transactions').innerHTML = 
                    `<div class="alert alert-danger">Σφάλμα φόρτωσης: ${error.message}</div>`;
            }
        }

        document.getElementById('form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const btn = e.target.querySelector('button');
            btn.disabled = true;
            btn.textContent = 'Προσθήκη...';
            
            try {
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
                if (!resp.ok) throw new Error(result.error || 'Σφάλμα');
                
                showMessage('✅ Η συναλλαγή προστέθηκε!', 'success');
                document.getElementById('desc').value = '';
                document.getElementById('amount').value = '';
                await loadData();
            } catch (error) {
                showMessage('❌ ' + error.message, 'danger');
            } finally {
                btn.disabled = false;
                btn.textContent = 'Προσθήκη';
            }
        });

        async function deleteTransaction(id) {
            if (!confirm('Είστε σίγουροι ότι θέλετε να διαγράψετε αυτή τη συναλλαγή;')) return;
            try {
                await fetch('/api/del/' + id, {method: 'DELETE'});
                showMessage('🗑️ Η συναλλαγή διαγράφηκε!', 'success');
                await loadData();
            } catch (error) {
                showMessage('❌ Σφάλμα διαγραφής', 'danger');
            }
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
            
            // Δείχνουμε μήνυμα ότι το Excel θα κατέβει
            showMessage('📥 Δημιουργία Excel... Θα κατέβει αυτόματα!', 'info');
            
            // Κατέβασμα αρχείου
            window.location.href = url;
        }

        async function importExcel(input) {
            const file = input.files[0];
            if (!file) return;
            
            const formData = new FormData();
            formData.append('file', file);
            
            try {
                const resp = await fetch('/api/import', {
                    method: 'POST', 
                    body: formData
                });
                
                const result = await resp.json();
                if (!resp.ok) throw new Error(result.error);
                
                showMessage(result.message || 'Επιτυχής εισαγωγή!', 'success');
                await loadData();
            } catch (error) {
                showMessage('❌ Σφάλμα εισαγωγής: ' + error.message, 'danger');
            }
            
            input.value = '';
        }

        // Φόρτωση δεδομένων κατά την εκκίνηση
        loadData();
        
        // Αυτόματη ανανέωση κάθε 30 δευτερόλεπτα
        setInterval(loadData, 30000);
    </script>
</body>
</html>
'''

# ─── ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ ───
def init_db():
    try:
        conn = sqlite3.connect(str(DB_FILE))
        conn.execute('''CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            description TEXT NOT NULL,
            category TEXT DEFAULT 'Γενικά',
            amount REAL NOT NULL)''')
        
        # Έλεγχος για στήλη category (παλιές βάσεις)
        cursor = conn.execute("PRAGMA table_info(transactions)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'category' not in columns:
            conn.execute("ALTER TABLE transactions ADD COLUMN category TEXT DEFAULT 'Γενικά'")
        
        # Μετατροπή παλιών ημερομηνιών
        rows = conn.execute("SELECT id, date FROM transactions WHERE date LIKE '%/%'").fetchall()
        for row in rows:
            parts = row[1].split('/')
            if len(parts) == 3:
                conn.execute("UPDATE transactions SET date = ? WHERE id = ?", 
                           (f"{parts[2]}-{parts[1]}-{parts[0]}", row[0]))
        
        conn.commit()
        conn.close()
        print("✓ Βάση δεδομένων έτοιμη")
        return True
    except Exception as e:
        print(f"❌ Σφάλμα βάσης: {e}")
        traceback.print_exc()
        return False

# ─── ROUTES ───
@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/api/data')
def get_data():
    try:
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
        
        # Στατιστικά
        if start and end:
            balance = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE date BETWEEN ? AND ?", 
                [start, end]
            ).fetchone()[0]
            income = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE amount > 0 AND date BETWEEN ? AND ?", 
                [start, end]
            ).fetchone()[0]
            expenses = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE amount < 0 AND date BETWEEN ? AND ?", 
                [start, end]
            ).fetchone()[0]
        else:
            balance = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM transactions").fetchone()[0]
            income = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE amount > 0").fetchone()[0]
            expenses = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE amount < 0").fetchone()[0]
        
        conn.close()
        
        return jsonify({
            'transactions': transactions,
            'balance': float(balance),
            'income': float(income),
            'expenses': float(expenses)
        })
    except Exception as e:
        print(f"❌ Σφάλμα API: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/add', methods=['POST'])
def add():
    try:
        data = request.get_json(force=True)
        
        date = data.get('date', '')
        description = data.get('description', '')
        category = data.get('category', 'Γενικά')
        
        try:
            amount = float(data.get('amount', 0))
        except (ValueError, TypeError):
            return jsonify({'error': 'Μη έγκυρο ποσό'}), 400
        
        if not date or not description:
            return jsonify({'error': 'Συμπληρώστε όλα τα πεδία'}), 400
        
        conn = sqlite3.connect(str(DB_FILE))
        cursor = conn.execute(
            "INSERT INTO transactions (date, description, category, amount) VALUES (?, ?, ?, ?)",
            (date, description, category, amount)
        )
        conn.commit()
        new_id = cursor.lastrowid
        conn.close()
        
        print(f"✓ Συναλλαγή #{new_id} προστέθηκε")
        return jsonify({'success': True, 'id': new_id}), 201
        
    except Exception as e:
        print(f"❌ Σφάλμα προσθήκης: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/del/<int:id>', methods=['DELETE'])
def delete(id):
    try:
        conn = sqlite3.connect(str(DB_FILE))
        conn.execute("DELETE FROM transactions WHERE id = ?", (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/export')
def export():
    try:
        conn = sqlite3.connect(str(DB_FILE))
        start = request.args.get('start_date')
        end = request.args.get('end_date')
        
        query = """
            SELECT date as 'Ημερομηνία', description as 'Περιγραφή', 
                   category as 'Κατηγορία', amount as 'Ποσό' 
            FROM transactions
        """
        params = []
        if start and end:
            query += " WHERE date BETWEEN ? AND ?"
            params = [start, end]
        query += " ORDER BY date DESC"
        
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        
        # Μετατροπή ημερομηνιών σε DD/MM/YYYY
        if 'Ημερομηνία' in df.columns:
            df['Ημερομηνία'] = df['Ημερομηνία'].apply(
                lambda d: f"{d[8:10]}/{d[5:7]}/{d[:4]}" 
                if isinstance(d, str) and len(d) == 10 and d[4] == '-' else d
            )
        
        # Αποθήκευση και στον τοπικό φάκελο
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        local_path = EXPORTS_FOLDER / f'oikonomika_{timestamp}.xlsx'
        
        with pd.ExcelWriter(local_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Οικονομικά')
        
        print(f"✓ Excel αποθηκεύτηκε: {local_path}")
        
        # Αποστολή για download
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'oikonomika_{timestamp}.xlsx'
        )
    except Exception as e:
        print(f"❌ Σφάλμα εξαγωγής: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/import', methods=['POST'])
def import_excel():
    try:
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
                date_val = str(row['Ημερομηνία'])
                # Μετατροπή από DD/MM/YYYY σε YYYY-MM-DD
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
            except Exception as e:
                print(f"⚠️ Παράλειψη γραμμής: {e}")
                continue
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'success': True,
            'message': f'✅ Εισήχθησαν {imported} εγγραφές με επιτυχία!'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ─── ΕΚΚΙΝΗΣΗ ───
if __name__ == '__main__':
    print("\n🔧 Αρχικοποίηση...")
    
    if not init_db():
        print("❌ Σφάλμα στη βάση δεδομένων")
        input("Πατήστε Enter για έξοδο...")
        sys.exit(1)
    
    port = 5000
    url = f'http://127.0.0.1:{port}'
    
    print(f"\n{'='*60}")
    print(f"  ✅ Η εφαρμογή είναι έτοιμη!")
    print(f"  🌐 Διεύθυνση: {url}")
    print(f"  📂 Δεδομένα: {DB_FILE}")
    print(f"  📁 Excel: {EXPORTS_FOLDER}")
    print(f"  💡 Μην κλείσετε αυτό το παράθυρο όσο χρησιμοποιείτε την εφαρμογή")
    print(f"{'='*60}\n")
    
    # Άνοιγμα browser
    def open_browser():
        time.sleep(2)
        webbrowser.open(url)
    
    threading.Thread(target=open_browser, daemon=True).start()
    
    # Εκκίνηση server
    try:
        app.run(host='127.0.0.1', port=port, debug=False)
    except KeyboardInterrupt:
        print("\n👋 Η εφαρμογή τερματίστηκε.")
    except Exception as e:
        print(f"\n❌ Σφάλμα: {e}")
        traceback.print_exc()
        input("Πατήστε Enter για έξοδο...")
