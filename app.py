from flask import Flask, render_template, request, send_file, session
import pandas as pd
from fuzzywuzzy import fuzz
import os
import pickle
import uuid

app = Flask(__name__)
app.secret_key = 'supersecretkey'


# ==================== FITUR 1: BANDINGKAN EXCEL ====================
def compare_sheets(sheet_a, sheet_b):
    result = []
    used_a = set()
    for _, row_b in sheet_b.iterrows():
        match = None
        method = ""
        similarity = ""

        for idx_a, row_a in sheet_a.iterrows():
            if idx_a in used_a:
                continue
            if row_a.get("UUID") == row_b.get("UUID") and pd.notna(row_b.get("UUID")):
                match = row_a
                method = "UUID"
                break
            elif row_a.get("ISBN") == row_b.get("ISBN") and pd.notna(row_b.get("ISBN")):
                match = row_a
                method = "ISBN"
                break
            elif row_a.get("EISBN") == row_b.get("EISBN") and pd.notna(row_b.get("EISBN")):
                match = row_a
                method = "EISBN"
                break
            elif pd.notna(row_b.get("Judul")) and pd.notna(row_a.get("Judul")):
                score = fuzz.token_sort_ratio(str(row_a['Judul']), str(row_b['Judul']))
                if score >= 85:
                    match = row_a
                    method = "Judul (Fuzzy)"
                    similarity = score
                    break

        if match is not None:
            used_a.add(match.name)
            result.append({
                "UUID_A": match.get("UUID"),
                "UUID_B": row_b.get("UUID"),
                "Judul_A": match.get("Judul"),
                "Judul_B": row_b.get("Judul"),
                "Metode": method,
                "Similarity": similarity
            })
        else:
            result.append({
                "UUID_A": None,
                "UUID_B": row_b.get("UUID"),
                "Judul_A": None,
                "Judul_B": row_b.get("Judul"),
                "Metode": "Tidak ditemukan",
                "Similarity": ""
            })
    return result


def compare_excels(file_a, file_b, mode='multi', sheet_a_name=None, sheet_b_name=None):
    try:
        df_a = pd.read_excel(file_a, sheet_name=None)
        df_b = pd.read_excel(file_b, sheet_name=None)
    except Exception as e:
        return {"Error": [{"Metode": "Gagal membaca file Excel"}]}

    results_by_sheet = {}

    if mode == "single":
        if not sheet_a_name or not sheet_b_name:
            return {"Error": [{"Metode": "Nama sheet tidak lengkap untuk mode single"}]}

        sheet_a = df_a.get(sheet_a_name)
        sheet_b = df_b.get(sheet_b_name)

        if sheet_a is None or sheet_b is None:
            return {"Error": [{"Metode": f"Sheet tidak ditemukan: {sheet_a_name} atau {sheet_b_name}"}]}

        results_by_sheet[f"{sheet_a_name} vs {sheet_b_name}"] = compare_sheets(sheet_a, sheet_b)
    else:
        try:
            sheet_a = list(df_a.values())[0]
        except IndexError:
            return {"Error": [{"Metode": "File A tidak memiliki sheet"}]}

        for name, sheet_b in df_b.items():
            results_by_sheet[name] = compare_sheets(sheet_a, sheet_b)

    return results_by_sheet


# ==================== FITUR 2: FILTER KATALOG ====================
def filter_excel_by_criteria(file, referensi_filter, kategori_filter, harga_min, harga_max, tahun_filter):
    xls = pd.ExcelFile(file)
    filtered_results = {}

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)

            # Normalisasi kolom
            df.columns = df.columns.str.strip()

            # Cari kolom tahun yang cocok
            tahun_col = next((col for col in df.columns if 'Tahun' in col and 'Digital' in col), None)

            if tahun_col and {'Referensi', 'Kategori*', 'HARGA SATUAN'}.issubset(df.columns):
                df['Referensi'] = df['Referensi'].astype(str).str.strip().str.lower()
                df['Kategori*'] = df['Kategori*'].astype(str).str.strip()
                df['HARGA SATUAN'] = pd.to_numeric(df['HARGA SATUAN'], errors='coerce')
                df[tahun_col] = pd.to_numeric(df[tahun_col], errors='coerce')

                # Terapkan filter jika ada input
                if referensi_filter:
                    referensi_filter = referensi_filter.lower().strip()
                    df = df[df['Referensi'].str.contains(referensi_filter, na=False)]

                if kategori_filter:
                    kategori_filter = kategori_filter.strip()
                    df = df[df['Kategori*'] == kategori_filter]

                if harga_min is not None:
                    df = df[df['HARGA SATUAN'] >= harga_min]

                if harga_max is not None:
                    df = df[df['HARGA SATUAN'] <= harga_max]

                if tahun_filter:
                    df = df[df[tahun_col].isin(tahun_filter)]

                # Simpan hasil jika tidak kosong
                if not df.empty:
                    filtered_results[sheet[:31]] = df

        except Exception as e:
            print(f"Gagal memproses sheet {sheet}: {e}")
            continue

    return filtered_results




# ==================== ROUTES ====================
@app.route('/')
def home():
    return render_template('index.html')


@app.route('/compare', methods=['GET', 'POST'])
def compare():
    if request.method == 'POST':
        file_a = request.files['fileA']
        file_b = request.files['fileB']
        mode = request.form.get('mode')
        sheet_a_name = request.form.get('sheetA')
        sheet_b_name = request.form.get('sheetB')

        results = compare_excels(file_a, file_b, mode, sheet_a_name, sheet_b_name)

        session_id = str(uuid.uuid4())
        os.makedirs("tmp", exist_ok=True)
        with open(f"tmp/{session_id}.pkl", "wb") as f:
            pickle.dump(results, f)
        session["comparison_file"] = session_id

        return render_template('result.html', results=results)
    return render_template('compare.html')


@app.route('/filter', methods=['GET', 'POST'])
def filter():
    if request.method == 'POST':
        file = request.files['file']
        referensi = request.form.get('referensi', '').strip()
        kategori = request.form.get('kategori', '').strip()
        harga_min_raw = request.form.get('harga_min', '').strip()
        harga_max_raw = request.form.get('harga_max', '').strip()
        tahun_raw = request.form.get('tahun_filter', '').strip()

        try:
            harga_min = float(harga_min_raw) if harga_min_raw else None
            harga_max = float(harga_max_raw) if harga_max_raw else None
        except ValueError:
            return "Input harga tidak valid."

        try:
            tahun_filter = [int(t.strip()) for t in tahun_raw.split(',') if t.strip()] if tahun_raw else None
        except ValueError:
            return "Input tahun tidak valid."

        results = filter_excel_by_criteria(file, referensi, kategori, harga_min, harga_max, tahun_filter)

        session_id = str(uuid.uuid4())
        os.makedirs("tmp", exist_ok=True)
        with open(f"tmp/{session_id}_filter.pkl", "wb") as f:
            pickle.dump(results, f)
        session["filter_file"] = session_id

        return render_template('filter_result.html', results=results)
    return render_template('filter.html')



@app.route('/export')
def export():
    session_id = session.get("comparison_file")
    if not session_id:
        return "Tidak ada data untuk diekspor."

    pkl_path = f"tmp/{session_id}.pkl"
    if not os.path.exists(pkl_path):
        return "Tidak ada data untuk diekspor."

    with open(pkl_path, "rb") as f:
        results = pickle.load(f)

    filename = "hasil_perbandingan_export.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for sheet_name, data in results.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return send_file(filename, as_attachment=True)


@app.route('/export-filter')
def export_filter():
    session_id = session.get("filter_file")
    if not session_id:
        return "Tidak ada data filter untuk diekspor."

    pkl_path = f"tmp/{session_id}_filter.pkl"
    if not os.path.exists(pkl_path):
        return "Tidak ada data filter untuk diekspor."

    with open(pkl_path, "rb") as f:
        results = pickle.load(f)

    filename = "hasil_filter_export.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return send_file(filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
