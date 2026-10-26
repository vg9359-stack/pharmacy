import os
import sqlite3
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import openpyxl

app = Flask(__name__)
DB_NAME = "pharmacy.db"
UPLOAD_FOLDER = os.path.join("static", "uploads")
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "pdf"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Medicines Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS medicines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            batch_no TEXT NOT NULL,
            category TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            price REAL NOT NULL,
            expiry_date TEXT NOT NULL,
            storage_condition TEXT DEFAULT 'Room Temperature (<25°C)'
        )
    ''')
    
    # Sales Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            medicine_name TEXT NOT NULL,
            batch_no TEXT NOT NULL,
            customer_name TEXT,
            customer_phone TEXT,
            quantity INTEGER NOT NULL,
            unit_price REAL NOT NULL,
            discount_percent REAL DEFAULT 0,
            total_price REAL NOT NULL,
            prescription_file TEXT,
            sale_date TEXT NOT NULL,
            refill_due_date TEXT
        )
    ''')
    conn.commit()

    # --- SEED 50 DUMMY MEDICINES & 25 SALES IF ERASED/EMPTY ---
    cursor.execute("SELECT COUNT(*) FROM medicines")
    if cursor.fetchone()[0] == 0:
        sample_medicines = [
            ("Paracetamol 500mg", "BATCH-101", "Analgesic", 150, 45.00, "2027-12-31", "Room Temperature (<25°C)"),
            ("Amoxicillin 250mg", "BATCH-102", "Antibiotic", 45, 120.00, "2026-11-15", "Room Temperature (<25°C)"),
            ("Ibuprofen 400mg", "BATCH-103", "Anti-inflammatory", 80, 65.00, "2027-10-20", "Room Temperature (<25°C)"),
            ("Cetirizine 10mg", "BATCH-104", "Antihistamine", 120, 35.00, "2028-01-10", "Room Temperature (<25°C)"),
            ("Omeprazole 20mg", "BATCH-105", "Gastrointestinal", 8, 150.00, "2026-11-05", "Cool & Dry Place (<20°C)"),
            ("Azithromycin 500mg", "BATCH-106", "Antibiotic", 5, 180.00, "2027-04-12", "Protect from Light"),
            ("Metformin 500mg", "BATCH-107", "Antidiabetic", 200, 90.00, "2028-05-18", "Room Temperature (<25°C)"),
            ("Amlodipine 5mg", "BATCH-108", "Cardiovascular", 90, 110.00, "2027-11-25", "Room Temperature (<25°C)"),
            ("Insulin Glargine", "BATCH-109", "Antidiabetic", 12, 450.00, "2027-08-30", "Refrigerated (2°C–8°C)"),
            ("Covaxin Vaccine", "BATCH-110", "Vaccines", 6, 850.00, "2026-11-20", "Refrigerated (2°C–8°C)"),
            ("Loratadine 10mg", "BATCH-111", "Antihistamine", 14, 55.00, "2027-06-15", "Room Temperature (<25°C)"),
            ("Pantoprazole 40mg", "BATCH-112", "Gastrointestinal", 60, 140.00, "2027-09-10", "Room Temperature (<25°C)"),
            ("Losartan 50mg", "BATCH-113", "Cardiovascular", 75, 138.00, "2027-12-01", "Room Temperature (<25°C)"),
            ("Metoprolol 25mg", "BATCH-114", "Cardiovascular", 4, 105.00, "2026-11-10", "Room Temperature (<25°C)"),
            ("Vitamin C 1000mg", "BATCH-115", "Supplements", 300, 50.00, "2028-06-30", "Room Temperature (<25°C)"),
            ("Vitamin D3 60K", "BATCH-116", "Supplements", 180, 85.00, "2028-04-15", "Protect from Light"),
            ("Levothyroxine 50mcg", "BATCH-117", "Endocrine", 95, 160.00, "2027-11-10", "Cool & Dry Place (<20°C)"),
            ("Ciprofloxacin 500mg", "BATCH-118", "Antibiotic", 40, 195.00, "2027-10-05", "Room Temperature (<25°C)"),
            ("Ranitidine 150mg", "BATCH-119", "Gastrointestinal", 3, 60.00, "2026-11-02", "Room Temperature (<25°C)"),
            ("Dexamethasone 4mg", "BATCH-120", "Steroid", 55, 128.00, "2027-08-19", "Protect from Light"),
            ("Clopidogrel 75mg", "BATCH-121", "Cardiovascular", 70, 250.00, "2028-02-28", "Room Temperature (<25°C)"),
            ("Aspirin 81mg", "BATCH-122", "Analgesic", 250, 35.00, "2028-07-15", "Room Temperature (<25°C)"),
            ("Gabapentin 300mg", "BATCH-123", "Neurology", 9, 280.00, "2027-09-30", "Room Temperature (<25°C)"),
            ("Tramadol 50mg", "BATCH-124", "Analgesic", 2, 175.00, "2026-11-08", "Room Temperature (<25°C)"),
            ("Montelukast 10mg", "BATCH-125", "Respiratory", 85, 210.00, "2028-01-25", "Room Temperature (<25°C)"),
            ("Salbutamol Inhaler", "BATCH-126", "Respiratory", 60, 145.00, "2027-12-15", "Room Temperature (<25°C)"),
            ("Atorvastatin 10mg", "BATCH-127", "Cardiovascular", 110, 225.00, "2028-03-14", "Room Temperature (<25°C)"),
            ("Telmisartan 40mg", "BATCH-128", "Cardiovascular", 85, 130.00, "2027-10-18", "Room Temperature (<25°C)"),
            ("Glimepiride 2mg", "BATCH-129", "Antidiabetic", 130, 95.00, "2028-02-10", "Room Temperature (<25°C)"),
            ("Sitagliptin 100mg", "BATCH-130", "Antidiabetic", 40, 310.00, "2027-11-01", "Room Temperature (<25°C)"),
            ("Doxycycline 100mg", "BATCH-131", "Antibiotic", 65, 115.00, "2027-07-19", "Protect from Light"),
            ("Ofloxacin 200mg", "BATCH-132", "Antibiotic", 50, 88.00, "2027-09-09", "Room Temperature (<25°C)"),
            ("Fluconazole 150mg", "BATCH-133", "Antifungal", 75, 42.00, "2028-04-10", "Room Temperature (<25°C)"),
            ("Acyclovir 400mg", "BATCH-134", "Antiviral", 35, 165.00, "2027-08-25", "Room Temperature (<25°C)"),
            ("Ondansetron 4mg", "BATCH-135", "Gastrointestinal", 110, 52.00, "2028-01-18", "Room Temperature (<25°C)"),
            ("Domperidone 10mg", "BATCH-136", "Gastrointestinal", 90, 48.00, "2027-12-05", "Room Temperature (<25°C)"),
            ("Loperamide 2mg", "BATCH-137", "Gastrointestinal", 140, 25.00, "2028-05-30", "Room Temperature (<25°C)"),
            ("Budesonide Inhaler", "BATCH-138", "Respiratory", 25, 380.00, "2027-11-22", "Room Temperature (<25°C)"),
            ("Levocetirizine 5mg", "BATCH-139", "Antihistamine", 160, 40.00, "2028-06-12", "Room Temperature (<25°C)"),
            ("Fexofenadine 120mg", "BATCH-140", "Antihistamine", 70, 110.00, "2027-10-30", "Room Temperature (<25°C)"),
            ("Prednisolone 10mg", "BATCH-141", "Steroid", 80, 62.00, "2027-09-14", "Protect from Light"),
            ("Methylprednisolone 8mg", "BATCH-142", "Steroid", 45, 140.00, "2027-08-01", "Protect from Light"),
            ("Calcium + Vit D3", "BATCH-143", "Supplements", 220, 120.00, "2028-07-20", "Room Temperature (<25°C)"),
            ("Zinc Sulphate 20mg", "BATCH-144", "Supplements", 190, 35.00, "2028-03-15", "Room Temperature (<25°C)"),
            ("B-Complex Capsules", "BATCH-145", "Supplements", 210, 45.00, "2028-08-10", "Room Temperature (<25°C)"),
            ("Iron + Folic Acid", "BATCH-146", "Supplements", 175, 68.00, "2028-02-28", "Room Temperature (<25°C)"),
            ("Pregabalin 75mg", "BATCH-147", "Neurology", 38, 240.00, "2027-11-15", "Room Temperature (<25°C)"),
            ("Escitalopram 10mg", "BATCH-148", "Neurology", 52, 135.00, "2027-12-20", "Room Temperature (<25°C)"),
            ("Alprazolam 0.25mg", "BATCH-149", "Neurology", 7, 30.00, "2026-11-01", "Protect from Light"),
            ("Clonazepam 0.5mg", "BATCH-150", "Neurology", 11, 42.00, "2027-05-10", "Protect from Light")
        ]
        cursor.executemany('''
            INSERT INTO medicines (name, batch_no, category, quantity, price, expiry_date, storage_condition)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', sample_medicines)

        sample_sales = [
            ("Paracetamol 500mg", "BATCH-101", "Rahul Sharma", "+919876543210", 10, 45.00, 0, 450.00, None, "2026-10-01 10:15:00", "2026-10-31"),
            ("Amoxicillin 250mg", "BATCH-102", "Priya Patel", "+919876543211", 5, 120.00, 5, 570.00, None, "2026-10-02 14:30:00", "2026-11-01"),
            ("Azithromycin 500mg", "BATCH-106", "Amit Verma", "+919876543212", 8, 180.00, 0, 1440.00, None, "2026-10-04 11:20:00", "2026-11-03"),
            ("Insulin Glargine", "BATCH-109", "Suresh Kumar", "+919876543213", 2, 450.00, 0, 900.00, None, "2026-10-05 16:00:00", "2026-11-04"),
            ("Metformin 500mg", "BATCH-107", "Ananya Roy", "+919876543214", 20, 90.00, 10, 1620.00, None, "2026-10-07 09:45:00", "2026-11-06"),
            ("Vitamin C 1000mg", "BATCH-115", "Vikram Singh", "+919876543215", 15, 50.00, 0, 750.00, None, "2026-10-09 12:10:00", "2026-11-08"),
            ("Amlodipine 5mg", "BATCH-108", "Deepak Gupta", "+919876543216", 10, 110.00, 0, 1100.00, None, "2026-10-10 15:20:00", "2026-11-09"),
            ("Salbutamol Inhaler", "BATCH-126", "Meena Joshi", "+919876543217", 3, 145.00, 0, 435.00, None, "2026-10-12 11:00:00", "2026-11-11"),
            ("Atorvastatin 10mg", "BATCH-127", "Rajesh Nair", "+919876543218", 12, 225.00, 5, 2565.00, None, "2026-10-14 17:30:00", "2026-11-13"),
            ("Pantoprazole 40mg", "BATCH-112", "Sunita Rao", "+919876543219", 10, 140.00, 0, 1400.00, None, "2026-10-15 13:15:00", "2026-11-14"),
            ("Cetirizine 10mg", "BATCH-104", "Karan Mehta", "+919876543220", 25, 35.00, 0, 875.00, None, "2026-10-17 10:00:00", "2026-11-16"),
            ("Ibuprofen 400mg", "BATCH-103", "Pooja Hegde", "+919876543221", 8, 65.00, 0, 520.00, None, "2026-10-18 14:10:00", "2026-11-17"),
            ("Budesonide Inhaler", "BATCH-138", "Arjun Das", "+919876543222", 2, 380.00, 0, 760.00, None, "2026-10-20 16:40:00", "2026-11-19"),
            ("Calcium + Vit D3", "BATCH-143", "Sneha Kapoor", "+919876543223", 15, 120.00, 5, 1710.00, None, "2026-10-21 11:25:00", "2026-11-20"),
            ("Gabapentin 300mg", "BATCH-123", "Manoj Tiwari", "+919876543224", 5, 280.00, 0, 1400.00, None, "2026-10-22 18:00:00", "2026-11-21"),
            ("Montelukast 10mg", "BATCH-125", "Kavita Reddy", "+919876543225", 10, 210.00, 0, 2100.00, None, "2026-10-23 10:45:00", "2026-11-22"),
            ("Telmisartan 40mg", "BATCH-128", "Sanjay Dutt", "+919876543226", 14, 130.00, 0, 1820.00, None, "2026-10-24 15:15:00", "2026-11-23"),
            ("Sitagliptin 100mg", "BATCH-130", "Geeta Phogat", "+919876543227", 6, 310.00, 0, 1860.00, None, "2026-10-25 09:30:00", "2026-11-24"),
            ("Omeprazole 20mg", "BATCH-105", "Rohan Prasad", "+919876543228", 4, 150.00, 0, 600.00, None, "2026-10-25 12:00:00", "2026-11-24"),
            ("Ranitidine 150mg", "BATCH-119", "Nisha Saxena", "+919876543229", 5, 60.00, 0, 300.00, None, "2026-10-25 14:15:00", "2026-11-24"),
            ("Pregabalin 75mg", "BATCH-147", "Harish Kumar", "+919876543230", 4, 240.00, 0, 960.00, None, "2026-10-26 10:00:00", "2026-11-25"),
            ("Fluconazole 150mg", "BATCH-133", "Swati Mishra", "+919876543231", 6, 42.00, 0, 252.00, None, "2026-10-26 11:30:00", "2026-11-25"),
            ("Levocetirizine 5mg", "BATCH-139", "Abhinav Bindra", "+919876543232", 20, 40.00, 0, 800.00, None, "2026-10-26 13:45:00", "2026-11-25"),
            ("Zinc Sulphate 20mg", "BATCH-144", "Divya Agarwal", "+919876543233", 10, 35.00, 0, 350.00, None, "2026-10-26 15:20:00", "2026-11-25"),
            ("Iron + Folic Acid", "BATCH-146", "Ritu Beri", "+919876543234", 12, 68.00, 0, 816.00, None, "2026-10-26 16:50:00", "2026-11-25")
        ]
        cursor.executemany('''
            INSERT INTO sales (medicine_name, batch_no, customer_name, customer_phone, quantity, unit_price, discount_percent, total_price, prescription_file, sale_date, refill_due_date)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', sample_sales)
        conn.commit()
    conn.close()

init_db()

# ------------------ ROUTES ------------------

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/dashboard", methods=["GET"])
def get_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM medicines ORDER BY expiry_date ASC")
    medicines = [dict(row) for row in cursor.fetchall()]

    cursor.execute("SELECT * FROM sales ORDER BY id DESC")
    sales = [dict(row) for row in cursor.fetchall()]

    cursor.execute("SELECT SUM(total_price) FROM sales")
    total_revenue = cursor.fetchone()[0] or 0.0

    today = datetime.now().date()
    thirty_days_later = today + timedelta(days=30)

    low_stock = [m for m in medicines if m["quantity"] < 10]
    expiring_soon = [
        m for m in medicines 
        if datetime.strptime(m["expiry_date"], "%Y-%m-%d").date() <= thirty_days_later
    ]

    # DEMAND FORECASTING & DAYS OF STOCK PREDICTOR
    forecasting_data = []
    for med in medicines:
        cursor.execute('''
            SELECT SUM(quantity) FROM sales 
            WHERE medicine_name = ? AND DATE(sale_date) >= DATE('now', '-30 days')
        ''', (med["name"],))
        sales_30_days = cursor.fetchone()[0] or 1
        daily_burn_rate = max(sales_30_days / 30.0, 0.1)
        days_left = round(med["quantity"] / daily_burn_rate, 1)

        forecasting_data.append({
            "name": med["name"],
            "batch_no": med["batch_no"],
            "current_stock": med["quantity"],
            "daily_burn_rate": round(daily_burn_rate, 2),
            "days_remaining": days_left,
            "status": "CRITICAL RUN OUT" if days_left < 7 else ("REORDER SOON" if days_left < 15 else "HEALTHY")
        })

    # REFILL ALERTS (Due within next 30 days)
    seven_days_later = (today + timedelta(days=30)).strftime("%Y-%m-%d")
    cursor.execute('''
        SELECT * FROM sales 
        WHERE refill_due_date IS NOT NULL AND refill_due_date <= ? AND refill_due_date >= ?
        ORDER BY refill_due_date ASC
    ''', (seven_days_later, today.strftime("%Y-%m-%d")))
    pending_refills = [dict(row) for row in cursor.fetchall()]

    # 5 TREND ANALYTICS QUERIES
    cursor.execute('''
        SELECT medicine_name, SUM(total_price) as total_rev 
        FROM sales GROUP BY medicine_name ORDER BY total_rev DESC LIMIT 5
    ''')
    top_revenue = [dict(row) for row in cursor.fetchall()]

    cursor.execute('SELECT category, SUM(quantity) as total_qty FROM medicines GROUP BY category')
    category_dist = [dict(row) for row in cursor.fetchall()]

    cursor.execute('SELECT DATE(sale_date) as date, SUM(total_price) as daily_rev FROM sales GROUP BY DATE(sale_date) ORDER BY date ASC')
    sales_timeline = [dict(row) for row in cursor.fetchall()]

    cursor.execute('SELECT name, quantity FROM medicines ORDER BY quantity ASC LIMIT 10')
    low_stock_trend = [dict(row) for row in cursor.fetchall()]

    cursor.execute('SELECT category, AVG(price) as avg_price FROM medicines GROUP BY category')
    avg_category_price = [dict(row) for row in cursor.fetchall()]

    conn.close()

    return jsonify({
        "medicines": medicines,
        "sales": sales,
        "metrics": {
            "total_medicines": len(medicines),
            "low_stock_count": len(low_stock),
            "expiring_soon_count": len(expiring_soon),
            "total_revenue": round(total_revenue, 2),
            "pending_refills_count": len(pending_refills)
        },
        "expiring_soon_items": expiring_soon,
        "forecasting": forecasting_data,
        "pending_refills": pending_refills,
        "trends": {
            "top_revenue": top_revenue,
            "category_dist": category_dist,
            "sales_timeline": sales_timeline,
            "low_stock_trend": low_stock_trend,
            "avg_category_price": avg_category_price
        }
    })

@app.route("/api/add_medicine", methods=["POST"])
def add_medicine():
    data = request.json
    try:
        conn = get_db_connection()
        conn.execute('''
            INSERT INTO medicines (name, batch_no, category, quantity, price, expiry_date, storage_condition)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            data["name"].strip(),
            data["batch_no"].strip(),
            data["category"].strip(),
            int(data["quantity"]),
            float(data["price"]),
            data["expiry_date"].strip(),
            data.get("storage_condition", "Room Temperature (<25°C)")
        ))
        conn.commit()
        conn.close()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

@app.route("/api/process_sale", methods=["POST"])
def process_sale():
    med_id = request.form.get("medicine_id")
    qty_requested = int(request.form.get("quantity", 0))
    customer_name = request.form.get("customer_name", "Walk-in Customer")
    customer_phone = request.form.get("customer_phone", "N/A")
    discount = float(request.form.get("discount", 0.0))

    file = request.files.get("prescription")
    filename = None
    if file and allowed_file(file.filename):
        sec_filename = secure_filename(file.filename)
        filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{sec_filename}"
        file.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM medicines WHERE id = ?", (med_id,))
    med = cursor.fetchone()

    if not med:
        conn.close()
        return jsonify({"success": False, "message": "Medicine not found."}), 404

    if qty_requested > med["quantity"]:
        conn.close()
        return jsonify({
            "success": False,
            "message": f"Insufficient stock! Available: {med['quantity']} units."
        }), 400

    new_qty = med["quantity"] - qty_requested
    subtotal = med["price"] * qty_requested
    total_price = round(subtotal * (1 - (discount / 100.0)), 2)
    now_dt = datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")

    refill_date = (now_dt + timedelta(days=30)).strftime("%Y-%m-%d")

    cursor.execute("UPDATE medicines SET quantity = ? WHERE id = ?", (new_qty, med_id))
    cursor.execute('''
        INSERT INTO sales (
            medicine_name, batch_no, customer_name, customer_phone,
            quantity, unit_price, discount_percent, total_price, prescription_file, sale_date, refill_due_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        med["name"], med["batch_no"], customer_name, customer_phone,
        qty_requested, med["price"], discount, total_price, filename, now_str, refill_date
    ))
    
    sale_id = cursor.lastrowid
    conn.commit()
    conn.close()

    whatsapp_msg = (
        f"Hello {customer_name}! 🏥 Thank you for shopping at PharmaCare Pro.\n"
        f"Receipt #INV-{sale_id}\n"
        f"Item: {med['name']} (Qty: {qty_requested})\n"
        f"Total Paid: INR {total_price:.2f}\n"
        f"Storage: {med['storage_condition']}\n"
        f"Next Refill Date: {refill_date}"
    )
    clean_phone = "".join(filter(str.isdigit, customer_phone))
    wa_url = f"https://api.whatsapp.com/send?phone={clean_phone}&text={whatsapp_msg.replace(' ', '%20').replace('\n', '%0A')}"

    return jsonify({"success": True, "sale_id": sale_id, "whatsapp_url": wa_url})

# ------------------ EXPORT PDF & EXCEL ------------------

@app.route("/api/export/pdf/<int:sale_id>", methods=["GET"])
def export_pdf_receipt(sale_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM sales WHERE id = ?", (sale_id,))
    sale = cursor.fetchone()
    conn.close()

    if not sale:
        return "Sale ID not found", 404

    sale_dict = dict(sale)
    pdf_path = f"receipt_{sale_id}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=letter)
    
    c.setFont("Helvetica-Bold", 18)
    c.drawString(100, 750, "PHARMACARE PRO RECEIPT")
    c.setFont("Helvetica", 10)
    c.drawString(100, 735, f"Date: {sale_dict.get('sale_date', 'N/A')} | Receipt #: INV-{sale_dict.get('id', sale_id)}")
    c.line(100, 725, 500, 725)

    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, 700, "Customer Information:")
    c.setFont("Helvetica", 10)
    c.drawString(100, 685, f"Name: {sale_dict.get('customer_name', 'Walk-in Customer')}")
    c.drawString(100, 670, f"Phone: {sale_dict.get('customer_phone', 'N/A')}")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, 640, "Order Summary:")
    c.setFont("Helvetica", 10)
    c.drawString(100, 620, f"Item: {sale_dict.get('medicine_name', 'N/A')} (Batch: {sale_dict.get('batch_no', 'N/A')})")
    c.drawString(100, 605, f"Quantity: {sale_dict.get('quantity', 0)} units")
    c.drawString(100, 590, f"Unit Price: INR {sale_dict.get('unit_price', 0.0):.2f}")
    c.drawString(100, 575, f"Discount: {sale_dict.get('discount_percent', 0.0)}%")
    c.line(100, 560, 500, 560)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, 540, f"Total Paid: INR {sale_dict.get('total_price', 0.0):.2f}")
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(100, 510, f"Next Refill Scheduled For: {sale_dict.get('refill_due_date', 'N/A')}")

    c.save()
    return send_file(pdf_path, as_attachment=True)

@app.route("/api/export/excel", methods=["GET"])
def export_excel():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM sales ORDER BY id DESC")
    sales = [dict(row) for row in cursor.fetchall()]
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Report"

    headers = ["Sale ID", "Medicine Name", "Batch No", "Customer Name", "Customer Phone", "Quantity", "Unit Price (INR)", "Discount %", "Total Price (INR)", "Refill Due Date", "Date"]
    ws.append(headers)

    for s in sales:
        ws.append([
            s.get("id"), s.get("medicine_name"), s.get("batch_no"), s.get("customer_name"),
            s.get("customer_phone"), s.get("quantity"), s.get("unit_price"), s.get("discount_percent"),
            s.get("total_price"), s.get("refill_due_date"), s.get("sale_date")
        ])

    excel_path = "Sales_Report.xlsx"
    wb.save(excel_path)
    return send_file(excel_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, port=5000)