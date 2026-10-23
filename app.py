import io
import base64
import openpyxl
from flask import Flask, render_template, request, jsonify

from processors import homi, website, outbound

app = Flask(__name__)
# Vercel membatasi body request ~4.5MB; lokal tetap boleh lebih besar
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max

# CATATAN: Tidak memakai cache di memory karena di serverless (Vercel)
# setiap request bisa ditangani instance berbeda. File Excel dikirim
# langsung (base64) bersama hasil preview.


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/inspect-sheets", methods=["POST"])
def inspect_sheets():
    file = request.files.get("file")
    if not file:
        return jsonify({"success": False, "error": "File tidak ditemukan."}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file.read()), read_only=True)
        sheets = wb.sheetnames
        wb.close()
        return jsonify({"success": True, "sheets": sheets})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/api/proses", methods=["POST"])
def proses():
    try:
        jenis = request.form.get("jenis", "")
        bulan = request.form.get("bulan", "")
        tahun = request.form.get("tahun", "")
        sheet_data = request.form.get("sheet_data", "").strip() or None
        sheet_sales = request.form.get("sheet_sales", "").strip() or None

        if not jenis or not bulan or not tahun:
            return jsonify({"error": "Jenis, Bulan, dan Tahun wajib diisi."}), 400

        tahun = int(tahun)

        if jenis == "homi":
            file = request.files.get("file_data")
            if not file:
                return jsonify({"error": "File Data HOMi wajib diupload."}), 400
            file_bytes = file.read()
            preview_rows, excel_bytes, filename = homi.process(
                file_bytes, bulan, tahun, sheet_name=sheet_data
            )
            kolom_labels = [
                "Bulan", "Tahun", "Nama Dealer", "Total Leads",
                "KACAB Belum FU", "KACAB %",
                "FLP Belum FU", "FLP %",
                "WL Target", "WL Actual",
                "Lap. Sopir", "CR Target", "CR Actual"
            ]

        elif jenis == "website":
            file = request.files.get("file_data")
            if not file:
                return jsonify({"error": "File Data Website wajib diupload."}), 400
            file_bytes = file.read()
            preview_rows, excel_bytes, filename = website.process(
                file_bytes, bulan, tahun, sheet_name=sheet_data
            )
            kolom_labels = [
                "Bulan", "Tahun", "Nama Dealer", "Total Leads",
                "KACAB Belum FU", "KACAB %",
                "FLP Belum FU", "FLP %",
                "WL Target", "WL Actual",
                "Lap. Sopir", "CR Target", "CR Actual"
            ]

        elif jenis == "outbound":
            file_data  = request.files.get("file_data")
            file_sales = request.files.get("file_sales")
            if not file_data or not file_sales:
                return jsonify({"error": "File Data Outbound dan File Sales wajib diupload."}), 400
            file_bytes  = file_data.read()
            sales_bytes = file_sales.read()
            preview_rows, excel_bytes, filename = outbound.process(
                file_bytes, sales_bytes, bulan, tahun,
                sheet_data=sheet_data, sheet_sales=sheet_sales
            )
            kolom_labels = [
                "Bulan", "Tahun", "Nama Dealer", "Total Leads",
                "KACAB Belum FU", "KACAB %",
                "FLP Belum FU", "FLP %",
                "WL Target", "WL Actual",
                "Lap. Sopir", "Contrib Target", "Contrib Actual",
                "Sales"
            ]
        else:
            return jsonify({"error": "Jenis laporan tidak dikenal."}), 400

        # Warna kolom
        if jenis in ("homi", "website"):
            warna_hdr  = homi.HTML_WARNA_HDR
            warna_data = homi.HTML_WARNA_DATA
        else:
            warna_hdr  = outbound.HTML_WARNA_HDR
            warna_data = outbound.HTML_WARNA_DATA

        return jsonify({
            "success": True,
            "filename": filename,
            "file_b64": base64.b64encode(excel_bytes).decode("ascii"),
            "kolom": kolom_labels,
            "rows": preview_rows,
            "warna_hdr": warna_hdr,
            "warna_data": warna_data,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("=" * 55)
    print("  >> Profiling System berjalan!")
    print("  >> Buka browser: http://localhost:5000")
    print("=" * 55)
    app.run(debug=True, port=5000)
