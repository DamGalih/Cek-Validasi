from flask import Flask, render_template, request, send_file, session
import pandas as pd
from rapidfuzz import fuzz
import os
import pickle
import uuid
from difflib import SequenceMatcher
from fuzzywuzzy import fuzz, process

app = Flask(__name__)
app.secret_key = 'supersecretkey'

def fuzzy_match(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ==================== FITUR 1: BANDINGKAN EXCEL (OPTIMASI CEPAT) ====================

def compare_sheets_fast(df1, df2):
    hasil = []

    # Pre-index untuk pencarian cepat
    df2_uuid = df2.set_index('UUID', drop=False) if 'UUID' in df2.columns else pd.DataFrame()
    df2_isbn = df2.set_index('ISBN', drop=False) if 'ISBN' in df2.columns else pd.DataFrame()
    df2_eisbn = df2.set_index('EISBN', drop=False) if 'EISBN' in df2.columns else pd.DataFrame()
    judul_list = df2['Judul'].dropna().astype(str).tolist() if 'Judul' in df2.columns else []

    for _, row1 in df1.iterrows():
        matched_row = None
        match_method = None

        uuid = row1.get('UUID')
        isbn = row1.get('ISBN')
        eisbn = row1.get('EISBN')
        judul = row1.get('Judul')

        # ===== 1. Cocok UUID =====
        if pd.notna(uuid) and uuid in df2_uuid.index:
            matched_row = df2_uuid.loc[uuid]
            match_method = 'UUID'

        # ===== 2. Cocok ISBN =====
        if matched_row is None and pd.notna(isbn) and isbn in df2_isbn.index:
            matched_row = df2_isbn.loc[isbn]
            match_method = 'ISBN'

        # ===== 3. Cocok EISBN =====
        if matched_row is None and pd.notna(eisbn) and eisbn in df2_eisbn.index:
            matched_row = df2_eisbn.loc[eisbn]
            match_method = 'EISBN'

        # ===== 4. Fuzzy Judul =====
        if matched_row is None and pd.notna(judul) and judul_list:
            best_match = process.extractOne(str(judul), judul_list, scorer=fuzz.ratio)
            if best_match and best_match[1] > 85:
                matched_row = df2[df2['Judul'] == best_match[0]].iloc[0]
                match_method = 'Judul (Fuzzy)'

        # ===== Simpan Hasil =====
        if matched_row is not None:
            harga1 = row1.get('Harga', 0) or 0
            harga2 = matched_row.get('Harga', 0) or 0
            selisih_harga = harga1 - harga2

            hasil.append({
                'Judul_Referensi': row1.get('Judul') or '-',
                'Judul_Katalog': matched_row.get('Judul') or '-',
                'UUID': uuid or matched_row.get('UUID') or '-',
                'ISBN': isbn or matched_row.get('ISBN') or '-',
                'EISBN': eisbn or matched_row.get('EISBN') or '-',
                'Harga_Referensi': harga1,
                'Harga_Katalog': harga2,
                'Selisih_Harga': selisih_harga,
                'Metode': match_method
            })

    return pd.DataFrame(hasil)


# ==================== PANGGIL FUNGSI BANDINGKAN ====================
def compare_excels(file_a, file_b, mode, sheet_a_name=None, sheet_b_name=None):
    xls_a = pd.ExcelFile(file_a)
    xls_b = pd.ExcelFile(file_b)
    results = {}

    if mode == "single":
        if not sheet_a_name or not sheet_b_name:
            raise ValueError("Nama Sheet A dan B wajib diisi untuk mode 'single'.")
        df_a = xls_a.parse(sheet_a_name)
        df_b = xls_b.parse(sheet_b_name)
        results["Perbandingan"] = compare_sheets_fast(df_a, df_b)

    elif mode == "multi":
        if not sheet_a_name:
            raise ValueError("Nama Sheet A wajib diisi untuk mode 'multi'.")
        df_a = xls_a.parse(sheet_a_name)
        for sheet in xls_b.sheet_names:
            df_b = xls_b.parse(sheet)
            results[f"{sheet_a_name} vs {sheet}"] = compare_sheets_fast(df_a, df_b)

    elif mode == "multi-matching":
        for sheet in set(xls_a.sheet_names).intersection(xls_b.sheet_names):
            df_a = xls_a.parse(sheet)
            df_b = xls_b.parse(sheet)
            results[sheet] = compare_sheets_fast(df_a, df_b)

    else:
        raise ValueError("Mode perbandingan tidak dikenali.")

    return results


# ==================== FITUR 2: FILTER KATALOG ====================
def filter_excel_by_criteria(file, referensi_filter, kategori_filter, harga_min, harga_max, tahun_filter):
    xls = pd.ExcelFile(file)
    filtered_results = {}

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
            df.columns = df.columns.str.strip()

            tahun_col = next((col for col in df.columns if 'Tahun' in col and 'Digital' in col), None)

            if tahun_col and {'Referensi', 'Kategori*', 'HARGA SATUAN'}.issubset(df.columns):
                df['Referensi'] = df['Referensi'].astype(str).str.strip().str.lower()
                df['Kategori*'] = df['Kategori*'].astype(str).str.strip()
                df['HARGA SATUAN'] = pd.to_numeric(df['HARGA SATUAN'], errors='coerce')
                df[tahun_col] = pd.to_numeric(df[tahun_col], errors='coerce')

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
        file_a = request.files.get('fileA')
        file_b = request.files.get('fileB')
        mode = request.form.get('mode')
        sheet_a_name = request.form.get("sheetA", "").strip()
        sheet_b_name = request.form.get("sheetB", "").strip()

        if not file_a or not file_b:
            return "File A dan File B wajib diunggah.", 400
        if mode == 'single' and (not sheet_a_name or not sheet_b_name):
            return "Nama Sheet A dan Sheet B wajib diisi untuk mode single.", 400

        results = compare_excels(file_a, file_b, mode, sheet_a_name, sheet_b_name)

        session_id = str(uuid.uuid4())
        os.makedirs("tmp", exist_ok=True)
        with open(f"tmp/{session_id}.pkl", "wb") as f:
            pickle.dump(results, f)
        session["comparison_file"] = session_id
        results = {k: v.to_dict(orient='records') for k, v in results.items()}
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

            # --- Filter hanya data yang ditemukan ---
            if "Keterangan" in df.columns:
                df = df[df["Keterangan"] != "Tidak ditemukan"]

            if not df.empty:
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
