"""
Flask Backend - Profiling System
"""
import base64
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


@app.route("/api/proses", methods=["POST"])
def proses():
    try:
        jenis = request.form.get("jenis", "")
        bulan = request.form.get("bulan", "")
        tahun = request.form.get("tahun", "")

        if not jenis or not bulan or not tahun:
            return jsonify({"error": "Jenis, Bulan, dan Tahun wajib diisi."}), 400

        tahun = int(tahun)

        if jenis == "homi":
            file = request.files.get("file_data")
            if not file:
                return jsonify({"error": "File Data HOMi wajib diupload."}), 400
            file_bytes = file.read()
            preview_rows, excel_bytes, filename = homi.process(file_bytes, bulan, tahun)
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
            preview_rows, excel_bytes, filename = website.process(file_bytes, bulan, tahun)
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
            preview_rows, excel_bytes, filename = outbound.process(file_bytes, sales_bytes, bulan, tahun)
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
