from flask import Flask, render_template, request, redirect, url_for, send_file
import csv
import os
from datetime import datetime

app = Flask(__name__)

# Vercel: hanya /tmp yang writable di serverless environment
DATA_DIR = "/tmp/wisuda_data"
CSV_FILE = os.path.join(DATA_DIR, "wisudawan.csv")

FIELDNAMES = [
    "id", "tanggal_daftar", "nama_lengkap", "nim",
    "fakultas", "program_studi", "nomor_hp", "email", "catatan"
]

def ensure_csv():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(CSV_FILE):
        with open(CSV_FILE, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(FIELDNAMES)


def read_data():
    ensure_csv()
    data = []
    with open(CSV_FILE, mode='r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            data.append(row)
    return data


def write_data(data):
    ensure_csv()
    with open(CSV_FILE, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(data)


@app.route("/")
def index():
    data = read_data()
    data = sorted(data, key=lambda x: int(x["id"]), reverse=True)
    return render_template("index.html", data=data)


@app.route("/tambah", methods=["GET", "POST"])
def tambah():
    if request.method == "POST":
        data = read_data()
        next_id = 1
        if data:
            next_id = max(int(row["id"]) for row in data) + 1

        new_row = {
            "id": str(next_id),
            "tanggal_daftar": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "nama_lengkap": request.form.get("nama_lengkap", "").strip(),
            "nim": request.form.get("nim", "").strip(),
            "fakultas": request.form.get("fakultas", "").strip(),
            "program_studi": request.form.get("program_studi", "").strip(),
            "nomor_hp": request.form.get("nomor_hp", "").strip(),
            "email": request.form.get("email", "").strip(),
            "catatan": request.form.get("catatan", "").strip(),
        }

        data.append(new_row)
        write_data(data)
        return redirect(url_for("index"))

    return render_template("tambah.html")


@app.route("/hapus/<int:id>")
def hapus(id):
    data = read_data()
    data = [row for row in data if int(row["id"]) != id]
    write_data(data)
    return redirect(url_for("index"))


@app.route("/export")
def export_csv():
    data = read_data()
    if not data:
        return redirect(url_for("index"))

    # Gunakan openpyxl langsung (tanpa pandas) agar lebih ringan
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Wisudawan"

    headers = ["ID", "Tanggal Daftar", "Nama Lengkap", "NIM",
               "Fakultas", "Program Studi", "Nomor HP", "Email", "Catatan"]

    header_fill = PatternFill("solid", fgColor="1A1208")
    header_font = Font(bold=True, color="D9B96A")
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for row in data:
        ws.append([
            row["id"], row["tanggal_daftar"], row["nama_lengkap"],
            row["nim"], row["fakultas"], row["program_studi"],
            row["nomor_hp"], row["email"], row["catatan"]
        ])

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 40)
        for cell in col_cells[1:]:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    filepath = os.path.join(DATA_DIR, "data_wisudawan.xlsx")
    wb.save(filepath)

    return send_file(
        filepath,
        as_attachment=True,
        download_name="data_wisudawan.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True)
