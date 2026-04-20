import os
import csv
import numpy as np
import pandas as pd
from scipy.stats import pearsonr
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# === Impor dari Modul Terpisah ===
from kalman_module import process_kalman
from kuantisasi_module import process_kuantisasi
try:
    from bch_module import process_bch
    from hash_module import process_hash
    from nist_module import process_nist
except ImportError:
    pass  # Akan dibuat nanti

# =====================================================================
# GLOBAL PARAMETERS
# =====================================================================
KUANTISASI_NUM_BITS = 3
BENCHMARK_ITERATIONS = 10
CHANPROB_TIME_SECONDS = 120.0  # Hardcoded sementara sesuai arahan: waktu channel probing

# Variasi Parameter Pengujian Skenario
PARAM_VARIATIONS = [
    {"q": 0.01, "r": 0.5, "bb": 1},
    {"q": 0.01, "r": 0.5, "bb": 5},
    {"q": 0.01, "r": 0.5, "bb": 50},
    {"q": 0.01, "r": 0.5, "bb": 100},
    {"q": 0.01, "r": 0.5, "bb": 200},
    {"q": 0.5, "r": 0.01, "bb": 1},
    {"q": 0.5, "r": 0.01, "bb": 5},
    {"q": 0.5, "r": 0.01, "bb": 50},
    {"q": 0.5, "r": 0.01, "bb": 100},
    {"q": 0.5, "r": 0.01, "bb": 200},
]

# Skenario iterasi 1 - 4
SCENARIOS = [1, 2, 3, 4]

# =====================================================================
# UTILITY FUNCTIONS
# =====================================================================
def read_rssi_csv(path):
    data = []
    if not os.path.exists(path):
        print(f"Warning: File {path} tidak ditemukan!")
        return data
        
    with open(path, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            try:
                data.append(int(row[0]))
            except:
                continue
    return data

def calculate_kdr(a, b):
    if not a or not b: return 0.0
    n = min(len(a), len(b))
    if n == 0: return 0.0
    diff = sum(1 for i in range(n) if a[i] != b[i])
    return (diff / n) * 100.0

def calc_corr(a, b):
    ln = min(len(a), len(b))
    if ln < 2: return "N/A"
    c, _ = pearsonr(a[:ln], b[:ln])
    return c

def calculate_cumulative_kgr(total_bits, *time_parts):
    """Hitung KGR kumulatif: total_bits / total_waktu_akumulasi."""
    total_time = 0.0
    for t in time_parts:
        try:
            total_time += float(t)
        except (TypeError, ValueError):
            continue
    if total_time <= 0:
        return 0.0
    return float(total_bits) / total_time

# =====================================================================
# EXCEL FORMATTING FUNCTIONS
# =====================================================================
def save_data_list(output_dir, filename, data_list, header):
    os.makedirs(output_dir, exist_ok=True)
    df = pd.DataFrame({header: data_list})
    df.to_excel(os.path.join(output_dir, filename), index=False)

def build_kalman_excel(output_path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Kalman"
    
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    current_row = 1
    
    for r in records:
        # Title Data Block
        title_val = f"Pengujian Skenario {r['skenario']} - Saat Q = {r['q']}, R = {r['r']}, BB = {r['bb']}"
        ws.cell(row=current_row, column=1, value=title_val).font = Font(bold=True, italic=True)
        current_row += 2
        
        start_row = current_row
        # Headers Top 
        ws.cell(row=start_row, column=1, value="Parameter").font = header_font
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+1, end_column=1)
        
        ws.cell(row=start_row, column=2, value="Sebelum Praproses").font = header_font
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=5)
        
        ws.cell(row=start_row, column=6, value="Setelah Praproses").font = header_font
        ws.merge_cells(start_row=start_row, start_column=6, end_row=start_row, end_column=9)
        
        # Headers Low
        cols_names = ["Alice", "Bob", "Eve-Alice", "Eve-Bob", "Alice", "Bob", "Eve-Alice", "Eve-Bob"]
        for idx, cname in enumerate(cols_names):
            ws.cell(row=start_row+1, column=2+idx, value=cname).font = header_font
            
        # Data Maksimum
        ws.cell(row=start_row+2, column=1, value="Maksimum (dBm)")
        vals_max = [r['orig_max_alice'], r['orig_max_bob'], r['orig_max_evealice'], r['orig_max_evebob'],
                    r['kalman_max_alice'], r['kalman_max_bob'], r['kalman_max_evealice'], r['kalman_max_evebob']]
        for idx, val in enumerate(vals_max): ws.cell(row=start_row+2, column=2+idx, value=val)
            
        # Data Minimum
        ws.cell(row=start_row+3, column=1, value="Minimum (dBm)")
        vals_min = [r['orig_min_alice'], r['orig_min_bob'], r['orig_min_evealice'], r['orig_min_evebob'],
                    r['kalman_min_alice'], r['kalman_min_bob'], r['kalman_min_evealice'], r['kalman_min_evebob']]
        for idx, val in enumerate(vals_min): ws.cell(row=start_row+3, column=2+idx, value=val)
            
        # Korelasi
        ws.cell(row=start_row+4, column=1, value="Koefisien Korelasi")
        
        # Sebelum Korelasi A&B + E&E (Merge 2 cell)
        c1 = ws.cell(row=start_row+4, column=2, value=r['orig_corr_ab'])
        c1.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=2, end_row=start_row+4, end_column=3)
        c2 = ws.cell(row=start_row+4, column=4, value=r['orig_corr_eve'])
        c2.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=4, end_row=start_row+4, end_column=5)
        
        # Sesudah Korelasi A&B + E&E
        c3 = ws.cell(row=start_row+4, column=6, value=r['kalman_corr_ab'])
        c3.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=6, end_row=start_row+4, end_column=7)
        c4 = ws.cell(row=start_row+4, column=8, value=r['kalman_corr_eve'])
        c4.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=8, end_row=start_row+4, end_column=9)
        
        # Waktu Komputasi
        ws.cell(row=start_row+5, column=1, value="Waktu Komputasi (s)")
        ws.merge_cells(start_row=start_row+5, start_column=1, end_row=start_row+5, end_column=5)
        for idx, val in enumerate([r['time_alice'], r['time_bob'], r['time_evealice'], r['time_evebob']]):
            ws.cell(row=start_row+5, column=6+idx, value=val)

        # Style alignment applying for all cells in block
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+5, min_col=1, max_col=9):
            for cell in row: cell.alignment = center_align

        # Additional 3 spaces for the next table
        current_row = start_row + 9 
        
    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 17
    ws.column_dimensions['A'].width = 25
        
    wb.save(output_path)

def build_kuantisasi_excel(output_path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Kuantisasi"
    
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    current_row = 1
    
    for r in records:
        title_val = f"Pengujian Skenario {r['skenario']} - Saat Q = {r['q']}, R = {r['r']}, BB = {r['bb']}"
        ws.cell(row=current_row, column=1, value=title_val).font = Font(bold=True, italic=True)
        current_row += 2
        
        start_row = current_row
        
        # Headers
        ws.cell(row=start_row, column=1, value="Parameter Performansi").font = header_font
        cols = ["Alice", "Bob", "Eve-Alice", "Eve-Bob"]
        for idx, val in enumerate(cols):
            ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        # KDR
        ws.cell(row=start_row+1, column=1, value="KDR (%)")
        ws.cell(row=start_row+1, column=2, value=r['kdr_ab'])
        ws.merge_cells(start_row=start_row+1, start_column=2, end_row=start_row+1, end_column=3)
        ws.cell(row=start_row+1, column=4, value=r['kdr_eve'])
        ws.merge_cells(start_row=start_row+1, start_column=4, end_row=start_row+1, end_column=5)
        
        # KGR
        ws.cell(row=start_row+2, column=1, value="KGR (bit/s)")
        for idx, val in enumerate([r['kgr_alice'], r['kgr_bob'], r['kgr_evealice'], r['kgr_evebob']]):
            ws.cell(row=start_row+2, column=2+idx, value=val)

        # Total bit yang dihasilkan
        ws.cell(row=start_row+3, column=1, value="Total Bit Dihasilkan")
        for idx, val in enumerate([r['total_bits_alice'], r['total_bits_bob'], r['total_bits_ea'], r['total_bits_eb']]):
            ws.cell(row=start_row+3, column=2+idx, value=val)
            
        # Waktu Komputasi
        ws.cell(row=start_row+4, column=1, value="Waktu komputasi (s)")
        for idx, val in enumerate([r['time_alice'], r['time_bob'], r['time_evealice'], r['time_evebob']]):
            ws.cell(row=start_row+4, column=2+idx, value=val)
            
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+4, min_col=1, max_col=5):
            for cell in row: cell.alignment = center_align

        current_row = start_row + 8
        
    for col in range(1, 6):
        ws.column_dimensions[get_column_letter(col)].width = 20
        
    wb.save(output_path)

def build_bch_excel(output_path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap BCH"
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    current_row = 1
    for r in records:
        ws.cell(row=current_row, column=1, value=f"Pengujian Skenario {r['skenario']} - Q={r['q']}, R={r['r']}, BB={r['bb']}").font = Font(bold=True, italic=True)
        current_row += 2
        start_row = current_row
        
        ws.cell(row=start_row, column=1, value="Parameter BCH").font = header_font
        for idx, val in enumerate(["A & B", "E-A & E-B"]):
            ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="KDR Setelah koreksi BCH (%)")
        ws.cell(row=start_row+1, column=2, value=r['kdr_after_ab'])
        ws.cell(row=start_row+1, column=3, value=r['kdr_after_eve'])
        
        ws.cell(row=start_row+2, column=1, value="KGR BCH Kumulatif (bit/s)")
        ws.cell(row=start_row+2, column=2, value=r['kgr_bch_ab'])
        ws.cell(row=start_row+2, column=3, value=r['kgr_bch_eve'])

        ws.cell(row=start_row+3, column=1, value="Parity Bits Dikirim")
        ws.cell(row=start_row+3, column=2, value=r['parity_bits_ab'])
        ws.cell(row=start_row+3, column=3, value=r['parity_bits_eve'])

        ws.cell(row=start_row+4, column=1, value="Total Bit (A/B | E-A/E-B)")
        ws.cell(row=start_row+4, column=2, value=f"{r['total_bits_alice']}/{r['total_bits_bob']}")
        ws.cell(row=start_row+4, column=3, value=f"{r['total_bits_ea']}/{r['total_bits_eb']}")

        ws.cell(row=start_row+5, column=1, value="Error Bit Sebelum")
        ws.cell(row=start_row+5, column=2, value=r['error_bits_ab_before'])
        ws.cell(row=start_row+5, column=3, value=r['error_bits_eve_before'])

        ws.cell(row=start_row+6, column=1, value="Error Bit Setelah")
        ws.cell(row=start_row+6, column=2, value=r['error_bits_ab_after'])
        ws.cell(row=start_row+6, column=3, value=r['error_bits_eve_after'])

        ws.cell(row=start_row+7, column=1, value="Bit Terkoreksi/Disamakan")
        ws.cell(row=start_row+7, column=2, value=r['corrected_bits_ab'])
        ws.cell(row=start_row+7, column=3, value=r['corrected_bits_eve'])

        ws.cell(row=start_row+8, column=1, value="Waktu Komputasi (s)")
        ws.cell(row=start_row+8, column=2, value=r['time_bch_ab'])
        ws.cell(row=start_row+8, column=3, value=r['time_bch_eve'])
        
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+8, min_col=1, max_col=3):
            for cell in row: cell.alignment = center_align
            
        current_row = start_row + 11
        
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 20
    wb.save(output_path)
def build_hash_excel(output_path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Hash"
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    current_row = 1
    for r in records:
        ws.cell(row=current_row, column=1, value=f"Skenario {r['skenario']} - Q={r['q']}, R={r['r']}, BB={r['bb']}").font = Font(bold=True, italic=True)
        current_row += 2
        start_row = current_row
        
        ws.cell(row=start_row, column=1, value="Parameter Hash").font = header_font
        cols = ["Alice", "Bob", "Eve-Alice", "Eve-Bob"]
        for idx, val in enumerate(cols):
            ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="Jumlah Kunci Cocok (Match)")
        ws.cell(row=start_row+1, column=2, value=r['aes_count_ab'])
        ws.merge_cells(start_row=start_row+1, start_column=2, end_row=start_row+1, end_column=3)
        ws.cell(row=start_row+1, column=4, value=r['aes_count_eve'])
        ws.merge_cells(start_row=start_row+1, start_column=4, end_row=start_row+1, end_column=5)

        ws.cell(row=start_row+2, column=1, value="Jumlah Kandidat Key")
        ws.cell(row=start_row+2, column=2, value=r['keys_count_alice'])
        ws.cell(row=start_row+2, column=3, value=r['keys_count_bob'])
        ws.cell(row=start_row+2, column=4, value=r['keys_count_ea'])
        ws.cell(row=start_row+2, column=5, value=r['keys_count_eb'])

        ws.cell(row=start_row+3, column=1, value="Total Bit Key (128*jumlah_key)")
        ws.cell(row=start_row+3, column=2, value=r['total_key_bits_alice'])
        ws.cell(row=start_row+3, column=3, value=r['total_key_bits_bob'])
        ws.cell(row=start_row+3, column=4, value=r['total_key_bits_ea'])
        ws.cell(row=start_row+3, column=5, value=r['total_key_bits_eb'])

        ws.cell(row=start_row+4, column=1, value="Total Bit AES Match")
        ws.cell(row=start_row+4, column=2, value=r['matched_key_bits_ab'])
        ws.merge_cells(start_row=start_row+4, start_column=2, end_row=start_row+4, end_column=3)
        ws.cell(row=start_row+4, column=4, value=r['matched_key_bits_eve'])
        ws.merge_cells(start_row=start_row+4, start_column=4, end_row=start_row+4, end_column=5)
        
        ws.cell(row=start_row+5, column=1, value="Kunci Pertama (Hex)")
        ws.cell(row=start_row+5, column=2, value=r['final_key_alice'])
        ws.cell(row=start_row+5, column=3, value=r['final_key_bob'])
        ws.cell(row=start_row+5, column=4, value=r['final_key_ea'])
        ws.cell(row=start_row+5, column=5, value=r['final_key_eb'])
        
        ws.cell(row=start_row+6, column=1, value="Waktu Komputasi (s)")
        ws.cell(row=start_row+6, column=2, value=r['time_hash_ab'])
        ws.cell(row=start_row+6, column=3, value=r['time_hash_ab'])
        ws.cell(row=start_row+6, column=4, value=r['time_hash_eve'])
        ws.cell(row=start_row+6, column=5, value=r['time_hash_eve'])
        
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+6, min_col=1, max_col=5):
            for cell in row:
                cell.alignment = center_align
        current_row = start_row + 9
        
    for col in range(1, 6): ws.column_dimensions[get_column_letter(col)].width = 38
    ws.column_dimensions['A'].width = 25
    wb.save(output_path)

def build_nist_excel(output_path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap NIST"
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    current_row = 1
    for r in records:
        ws.cell(row=current_row, column=1, value=f"Skenario {r['skenario']} - Q={r['q']}, R={r['r']}, BB={r['bb']}").font = Font(bold=True, italic=True)
        current_row += 2
        start_row = current_row
        
        ws.cell(row=start_row, column=1, value="Parameter NIST").font = header_font
        for idx, val in enumerate(["A & B", "E-A & E-B"]):
            ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="Jumlah Key Lulus")
        ws.cell(row=start_row+1, column=2, value=r['passed_keys_ab'])
        ws.cell(row=start_row+1, column=3, value=r['passed_keys_eve'])
        
        ws.cell(row=start_row+2, column=1, value="Rata-rata p-value (ApEn)")
        ws.cell(row=start_row+2, column=2, value=r['pval_ab'])
        ws.cell(row=start_row+2, column=3, value=r['pval_eve'])

        ws.cell(row=start_row+3, column=1, value="Pass Rate (%)")
        ws.cell(row=start_row+3, column=2, value=r['pass_rate_ab'])
        ws.cell(row=start_row+3, column=3, value=r['pass_rate_eve'])

        ws.cell(row=start_row+4, column=1, value="Distribusi p-value")
        ws.cell(row=start_row+4, column=2, value=r['pval_dist_ab'])
        ws.cell(row=start_row+4, column=3, value=r['pval_dist_eve'])
        
        ws.cell(row=start_row+5, column=1, value="Waktu Komputasi (s)")
        ws.cell(row=start_row+5, column=2, value=r['time_nist_ab'])
        ws.cell(row=start_row+5, column=3, value=r['time_nist_eve'])
        
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+5, min_col=1, max_col=3):
            for cell in row:
                cell.alignment = center_align
        current_row = start_row + 8
        
    for col in range(1, 4): ws.column_dimensions[get_column_letter(col)].width = 25
    wb.save(output_path)

def build_kalman_sheet(wb, records):
    ws = wb.create_sheet(title="Kalman")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    current_row = 1
    for r in records:
        title_val = f"Pengujian Skenario {r['skenario']} - Saat Q = {r['q']}, R = {r['r']}, BB = {r['bb']}"
        ws.cell(row=current_row, column=1, value=title_val).font = Font(bold=True, italic=True)
        current_row += 2
        
        start_row = current_row
        ws.cell(row=start_row, column=1, value="Parameter").font = header_font
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+1, end_column=1)
        
        ws.cell(row=start_row, column=2, value="Sebelum Praproses").font = header_font
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=5)
        
        ws.cell(row=start_row, column=6, value="Setelah Praproses").font = header_font
        ws.merge_cells(start_row=start_row, start_column=6, end_row=start_row, end_column=9)
        
        cols_names = ["Alice", "Bob", "Eve-Alice", "Eve-Bob", "Alice", "Bob", "Eve-Alice", "Eve-Bob"]
        for idx, cname in enumerate(cols_names):
            ws.cell(row=start_row+1, column=2+idx, value=cname).font = header_font
            
        ws.cell(row=start_row+2, column=1, value="Maksimum (dBm)")
        vals_max = [r['orig_max_alice'], r['orig_max_bob'], r['orig_max_evealice'], r['orig_max_evebob'], r['kalman_max_alice'], r['kalman_max_bob'], r['kalman_max_evealice'], r['kalman_max_evebob']]
        for idx, val in enumerate(vals_max): ws.cell(row=start_row+2, column=2+idx, value=val)
            
        ws.cell(row=start_row+3, column=1, value="Minimum (dBm)")
        vals_min = [r['orig_min_alice'], r['orig_min_bob'], r['orig_min_evealice'], r['orig_min_evebob'], r['kalman_min_alice'], r['kalman_min_bob'], r['kalman_min_evealice'], r['kalman_min_evebob']]
        for idx, val in enumerate(vals_min): ws.cell(row=start_row+3, column=2+idx, value=val)
            
        ws.cell(row=start_row+4, column=1, value="Koefisien Korelasi")
        c1 = ws.cell(row=start_row+4, column=2, value=r['orig_corr_ab'])
        c1.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=2, end_row=start_row+4, end_column=3)
        c2 = ws.cell(row=start_row+4, column=4, value=r['orig_corr_eve'])
        c2.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=4, end_row=start_row+4, end_column=5)
        c3 = ws.cell(row=start_row+4, column=6, value=r['kalman_corr_ab'])
        c3.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=6, end_row=start_row+4, end_column=7)
        c4 = ws.cell(row=start_row+4, column=8, value=r['kalman_corr_eve'])
        c4.number_format = '0.0000000000'
        ws.merge_cells(start_row=start_row+4, start_column=8, end_row=start_row+4, end_column=9)
        
        ws.cell(row=start_row+5, column=1, value="Waktu Komputasi (s)")
        ws.merge_cells(start_row=start_row+5, start_column=1, end_row=start_row+5, end_column=5)
        for idx, val in enumerate([r['time_alice'], r['time_bob'], r['time_evealice'], r['time_evebob']]): ws.cell(row=start_row+5, column=6+idx, value=val)

        for row in ws.iter_rows(min_row=start_row, max_row=start_row+5, min_col=1, max_col=9):
            for cell in row: cell.alignment = center_align

        current_row = start_row + 9 
        
    for col in range(1, 10): ws.column_dimensions[get_column_letter(col)].width = 17
    ws.column_dimensions['A'].width = 25

def build_kuantisasi_sheet(wb, records):
    ws = wb.create_sheet(title="Kuantisasi")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    current_row = 1
    for r in records:
        title_val = f"Pengujian Skenario {r['skenario']} - Saat Q = {r['q']}, R = {r['r']}, BB = {r['bb']}"
        ws.cell(row=current_row, column=1, value=title_val).font = Font(bold=True, italic=True)
        current_row += 2
        
        start_row = current_row
        ws.cell(row=start_row, column=1, value="Parameter Performansi").font = header_font
        cols = ["Alice", "Bob", "Eve-Alice", "Eve-Bob"]
        for idx, val in enumerate(cols): ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="KDR (%)")
        ws.cell(row=start_row+1, column=2, value=r['kdr_ab'])
        ws.merge_cells(start_row=start_row+1, start_column=2, end_row=start_row+1, end_column=3)
        ws.cell(row=start_row+1, column=4, value=r['kdr_eve'])
        ws.merge_cells(start_row=start_row+1, start_column=4, end_row=start_row+1, end_column=5)
        
        ws.cell(row=start_row+2, column=1, value="KGR (bit/s)")
        for idx, val in enumerate([r['kgr_alice'], r['kgr_bob'], r['kgr_evealice'], r['kgr_evebob']]): ws.cell(row=start_row+2, column=2+idx, value=val)

        ws.cell(row=start_row+3, column=1, value="Total Bit Dihasilkan")
        for idx, val in enumerate([r['total_bits_alice'], r['total_bits_bob'], r['total_bits_ea'], r['total_bits_eb']]): ws.cell(row=start_row+3, column=2+idx, value=val)
            
        ws.cell(row=start_row+4, column=1, value="Waktu komputasi (s)")
        for idx, val in enumerate([r['time_alice'], r['time_bob'], r['time_evealice'], r['time_evebob']]): ws.cell(row=start_row+4, column=2+idx, value=val)
            
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+4, min_col=1, max_col=5):
            for cell in row: cell.alignment = center_align

        current_row = start_row + 7
        
    for col in range(1, 6): ws.column_dimensions[get_column_letter(col)].width = 20

def build_bch_sheet(wb, records):
    ws = wb.create_sheet(title="BCH")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    current_row = 1
    for r in records:
        ws.cell(row=current_row, column=1, value=f"Pengujian Skenario {r['skenario']} - Q={r['q']}, R={r['r']}, BB={r['bb']}").font = Font(bold=True, italic=True)
        current_row += 2
        start_row = current_row
        
        ws.cell(row=start_row, column=1, value="Parameter BCH").font = header_font
        for idx, val in enumerate(["A & B", "E-A & E-B"]): ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="KDR Setelah koreksi BCH (%)")
        ws.cell(row=start_row+1, column=2, value=r['kdr_after_ab'])
        ws.cell(row=start_row+1, column=3, value=r['kdr_after_eve'])
        
        ws.cell(row=start_row+2, column=1, value="KGR BCH Kumulatif (bit/s)")
        ws.cell(row=start_row+2, column=2, value=r['kgr_bch_ab'])
        ws.cell(row=start_row+2, column=3, value=r['kgr_bch_eve'])

        ws.cell(row=start_row+3, column=1, value="Parity Bits Dikirim")
        ws.cell(row=start_row+3, column=2, value=r['parity_bits_ab'])
        ws.cell(row=start_row+3, column=3, value=r['parity_bits_eve'])

        ws.cell(row=start_row+4, column=1, value="Total Bit (A/B | E-A/E-B)")
        ws.cell(row=start_row+4, column=2, value=f"{r['total_bits_alice']}/{r['total_bits_bob']}")
        ws.cell(row=start_row+4, column=3, value=f"{r['total_bits_ea']}/{r['total_bits_eb']}")

        ws.cell(row=start_row+5, column=1, value="Error Bit Sebelum")
        ws.cell(row=start_row+5, column=2, value=r['error_bits_ab_before'])
        ws.cell(row=start_row+5, column=3, value=r['error_bits_eve_before'])

        ws.cell(row=start_row+6, column=1, value="Error Bit Setelah")
        ws.cell(row=start_row+6, column=2, value=r['error_bits_ab_after'])
        ws.cell(row=start_row+6, column=3, value=r['error_bits_eve_after'])

        ws.cell(row=start_row+7, column=1, value="Bit Terkoreksi/Disamakan")
        ws.cell(row=start_row+7, column=2, value=r['corrected_bits_ab'])
        ws.cell(row=start_row+7, column=3, value=r['corrected_bits_eve'])

        ws.cell(row=start_row+8, column=1, value="Waktu Komputasi (s)")
        ws.cell(row=start_row+8, column=2, value=r['time_bch_ab'])
        ws.cell(row=start_row+8, column=3, value=r['time_bch_eve'])
        
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+8, min_col=1, max_col=3):
            for cell in row: cell.alignment = center_align
            
        current_row = start_row + 11
        
    for col in range(1, 4): ws.column_dimensions[get_column_letter(col)].width = 25

def build_hash_sheet(wb, records):
    ws = wb.create_sheet(title="Hash_SHA_AES")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    current_row = 1
    for r in records:
        ws.cell(row=current_row, column=1, value=f"Skenario {r['skenario']} - Q={r['q']}, R={r['r']}, BB={r['bb']}").font = Font(bold=True, italic=True)
        current_row += 2
        start_row = current_row
        
        ws.cell(row=start_row, column=1, value="Parameter Hash").font = header_font
        cols = ["Alice", "Bob", "Eve-Alice", "Eve-Bob"]
        for idx, val in enumerate(cols): ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="Jumlah Kunci Cocok (Match)")
        ws.cell(row=start_row+1, column=2, value=r['aes_count_ab'])
        ws.merge_cells(start_row=start_row+1, start_column=2, end_row=start_row+1, end_column=3)
        ws.cell(row=start_row+1, column=4, value=r['aes_count_eve'])
        ws.merge_cells(start_row=start_row+1, start_column=4, end_row=start_row+1, end_column=5)

        ws.cell(row=start_row+2, column=1, value="Jumlah Kandidat Key")
        ws.cell(row=start_row+2, column=2, value=r['keys_count_alice'])
        ws.cell(row=start_row+2, column=3, value=r['keys_count_bob'])
        ws.cell(row=start_row+2, column=4, value=r['keys_count_ea'])
        ws.cell(row=start_row+2, column=5, value=r['keys_count_eb'])

        ws.cell(row=start_row+3, column=1, value="Total Bit Key (128*jumlah_key)")
        ws.cell(row=start_row+3, column=2, value=r['total_key_bits_alice'])
        ws.cell(row=start_row+3, column=3, value=r['total_key_bits_bob'])
        ws.cell(row=start_row+3, column=4, value=r['total_key_bits_ea'])
        ws.cell(row=start_row+3, column=5, value=r['total_key_bits_eb'])

        ws.cell(row=start_row+4, column=1, value="Total Bit AES Match")
        ws.cell(row=start_row+4, column=2, value=r['matched_key_bits_ab'])
        ws.merge_cells(start_row=start_row+4, start_column=2, end_row=start_row+4, end_column=3)
        ws.cell(row=start_row+4, column=4, value=r['matched_key_bits_eve'])
        ws.merge_cells(start_row=start_row+4, start_column=4, end_row=start_row+4, end_column=5)
        
        ws.cell(row=start_row+5, column=1, value="Kunci Pertama (Hex)")
        ws.cell(row=start_row+5, column=2, value=r['final_key_alice'])
        ws.cell(row=start_row+5, column=3, value=r['final_key_bob'])
        ws.cell(row=start_row+5, column=4, value=r['final_key_ea'])
        ws.cell(row=start_row+5, column=5, value=r['final_key_eb'])
        
        ws.cell(row=start_row+6, column=1, value="Waktu Komputasi (s)")
        ws.cell(row=start_row+6, column=2, value=r['time_hash_ab'])
        ws.cell(row=start_row+6, column=3, value=r['time_hash_ab'])
        ws.cell(row=start_row+6, column=4, value=r['time_hash_eve'])
        ws.cell(row=start_row+6, column=5, value=r['time_hash_eve'])
        
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+6, min_col=1, max_col=5):
            for cell in row: cell.alignment = center_align
        current_row = start_row + 9
        
    for col in range(1, 6): ws.column_dimensions[get_column_letter(col)].width = 38
    ws.column_dimensions['A'].width = 25

def build_nist_sheet(wb, records):
    ws = wb.create_sheet(title="NIST")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    current_row = 1
    for r in records:
        ws.cell(row=current_row, column=1, value=f"Skenario {r['skenario']} - Q={r['q']}, R={r['r']}, BB={r['bb']}").font = Font(bold=True, italic=True)
        current_row += 2
        start_row = current_row
        
        ws.cell(row=start_row, column=1, value="Parameter NIST").font = header_font
        for idx, val in enumerate(["A & B", "E-A & E-B"]): ws.cell(row=start_row, column=2+idx, value=val).font = header_font
            
        ws.cell(row=start_row+1, column=1, value="Jumlah Key Lulus")
        ws.cell(row=start_row+1, column=2, value=r['passed_keys_ab'])
        ws.cell(row=start_row+1, column=3, value=r['passed_keys_eve'])
        
        ws.cell(row=start_row+2, column=1, value="Rata-rata p-value (ApEn)")
        ws.cell(row=start_row+2, column=2, value=r['pval_ab'])
        ws.cell(row=start_row+2, column=3, value=r['pval_eve'])

        ws.cell(row=start_row+3, column=1, value="Pass Rate (%)")
        ws.cell(row=start_row+3, column=2, value=r['pass_rate_ab'])
        ws.cell(row=start_row+3, column=3, value=r['pass_rate_eve'])

        ws.cell(row=start_row+4, column=1, value="Distribusi p-value")
        ws.cell(row=start_row+4, column=2, value=r['pval_dist_ab'])
        ws.cell(row=start_row+4, column=3, value=r['pval_dist_eve'])
        
        ws.cell(row=start_row+5, column=1, value="Waktu Komputasi (s)")
        ws.cell(row=start_row+5, column=2, value=r['time_nist_ab'])
        ws.cell(row=start_row+5, column=3, value=r['time_nist_eve'])
        
        for row in ws.iter_rows(min_row=start_row, max_row=start_row+5, min_col=1, max_col=3):
            for cell in row: cell.alignment = center_align
        current_row = start_row + 8
        
    for col in range(1, 4): ws.column_dimensions[get_column_letter(col)].width = 25

# =====================================================================
# MAIN ENTRY POINT
# =====================================================================
def main():
    print("=== FULL SECRET KEY GENERATION (SKG) AUTOMATION ===")
    base_data = "data"
    output_base = "Output"
    
    # === Inisialisasi list untuk rekap global semua skenario ===
    global_kalman_records = []
    global_kuan_records = []
    global_bch_records = []
    global_hash_records = []
    global_nist_records = []
    
    for skenario in SCENARIOS:
        print(f"\n>>>> Memproses Skenario {skenario} <<<<")
        
        path_alice = os.path.join(base_data, "alice", f"skenario{skenario}_mita_alice.csv")
        path_bob   = os.path.join(base_data, "bob", f"skenario{skenario}_mita_bob.csv")
        path_eve_a = os.path.join(base_data, "eve alice", f"skenario{skenario}_mita_evealice.csv")
        path_eve_b = os.path.join(base_data, "eve bob", f"skenario{skenario}_mita_evebob.csv")
        
        raw_alice = read_rssi_csv(path_alice)
        raw_bob = read_rssi_csv(path_bob)
        raw_eve_a = read_rssi_csv(path_eve_a)
        raw_eve_b = read_rssi_csv(path_eve_b)
        
        if not (raw_alice and raw_bob and raw_eve_a and raw_eve_b):
            print(f"Melewati skenario {skenario} karena data file tidak lengkap di direktori.")
            continue
            
        skenario_out_dir = os.path.join(output_base, f"skenario_{skenario}")
        excel_kalman_dir = os.path.join(skenario_out_dir, "data_excel_kalman")
        excel_kuan_dir = os.path.join(skenario_out_dir, "data_excel_kuantisasi")
        os.makedirs(excel_kalman_dir, exist_ok=True)
        os.makedirs(excel_kuan_dir, exist_ok=True)
        
        kalman_records = []
        kuan_records = []
        bch_records = []
        hash_records = []
        nist_records = []
        
        for param in PARAM_VARIATIONS:
            q, r, bb = param['q'], param['r'], param['bb']
            print(f" -> Variasi dijalankan: Q={q}, R={r}, BB={bb}")
            
            total_len = min(len(raw_alice), len(raw_bob), len(raw_eve_a), len(raw_eve_b))
            cut_len = (total_len // bb) * bb

            # Pre-processing metrics should describe the full synchronized raw part,
            # not BB-dependent cropped data used by Kalman internals.
            ra_full = raw_alice[:total_len]
            rb_full = raw_bob[:total_len]
            rea_full = raw_eve_a[:total_len]
            reb_full = raw_eve_b[:total_len]

            # --- 1. Evaluasi Sebelum Praproses ---
            orig_max_alice = np.max(ra_full) if total_len > 0 else 0
            orig_max_bob = np.max(rb_full) if total_len > 0 else 0
            orig_max_evea = np.max(rea_full) if total_len > 0 else 0
            orig_max_eveb = np.max(reb_full) if total_len > 0 else 0

            orig_min_alice = np.min(ra_full) if total_len > 0 else 0
            orig_min_bob = np.min(rb_full) if total_len > 0 else 0
            orig_min_evea = np.min(rea_full) if total_len > 0 else 0
            orig_min_eveb = np.min(reb_full) if total_len > 0 else 0

            orig_corr_ab = calc_corr(ra_full, rb_full)
            orig_corr_eve = calc_corr(rea_full, reb_full)
            
            # --- 2. Filter Kalman Praproses (pindah ke module) ---
            kal_a, _kgr_kal_a_local, time_kal_a = process_kalman(raw_alice, q, r, bb, benchmark_iterations=BENCHMARK_ITERATIONS)
            kal_b, _kgr_kal_b_local, time_kal_b = process_kalman(raw_bob, q, r, bb, BENCHMARK_ITERATIONS)
            kal_ea, _kgr_kal_ea_local, time_kal_ea = process_kalman(raw_eve_a, q, r, bb, BENCHMARK_ITERATIONS)
            kal_eb, _kgr_kal_eb_local, time_kal_eb = process_kalman(raw_eve_b, q, r, bb, BENCHMARK_ITERATIONS)

            # Simpan Sinyal array ke excel (.xlsx)
            v_name = f"Q{q}_R{r}_BB{bb}"
            save_data_list(excel_kalman_dir, f"{v_name}_kalman_alice.xlsx", kal_a, "alice_kalman")
            save_data_list(excel_kalman_dir, f"{v_name}_kalman_bob.xlsx", kal_b, "bob_kalman")
            save_data_list(excel_kalman_dir, f"{v_name}_kalman_evealice.xlsx", kal_ea, "evealice_kalman")
            save_data_list(excel_kalman_dir, f"{v_name}_kalman_evebob.xlsx", kal_eb, "evebob_kalman")
            
            kal_max_alice = np.max(kal_a) if len(kal_a)>0 else 0
            kal_max_bob = np.max(kal_b) if len(kal_b)>0 else 0
            kal_max_evea = np.max(kal_ea) if len(kal_ea)>0 else 0
            kal_max_eveb = np.max(kal_eb) if len(kal_eb)>0 else 0
            
            kal_min_alice = np.min(kal_a) if len(kal_a)>0 else 0
            kal_min_bob = np.min(kal_b) if len(kal_b)>0 else 0
            kal_min_evea = np.min(kal_ea) if len(kal_ea)>0 else 0
            kal_min_eveb = np.min(kal_eb) if len(kal_eb)>0 else 0
            
            kal_corr_ab = calc_corr(kal_a, kal_b)
            kal_corr_eve = calc_corr(kal_ea, kal_eb)
            
            kalman_records.append({
                "skenario": skenario, "q": q, "r": r, "bb": bb,
                "orig_max_alice": orig_max_alice, "orig_max_bob": orig_max_bob, "orig_max_evealice": orig_max_evea, "orig_max_evebob": orig_max_eveb,
                "orig_min_alice": orig_min_alice, "orig_min_bob": orig_min_bob, "orig_min_evealice": orig_min_evea, "orig_min_evebob": orig_min_eveb,
                "orig_corr_ab": orig_corr_ab, "orig_corr_eve": orig_corr_eve,
                "kalman_max_alice": kal_max_alice, "kalman_max_bob": kal_max_bob, "kalman_max_evealice": kal_max_evea, "kalman_max_evebob": kal_max_eveb,
                "kalman_min_alice": kal_min_alice, "kalman_min_bob": kal_min_bob, "kalman_min_evealice": kal_min_evea, "kalman_min_evebob": kal_min_eveb,
                "kalman_corr_ab": kal_corr_ab, "kalman_corr_eve": kal_corr_eve,
                "time_alice": time_kal_a, "time_bob": time_kal_b, "time_evealice": time_kal_ea, "time_evebob": time_kal_eb
            })
            
            # --- 3. Kuantisasi Multibit 10x Iterasi (Pindah ke module) ---
            bs_a, _kgr_kuan_a_local, time_kuan_a = process_kuantisasi(kal_a, KUANTISASI_NUM_BITS, BENCHMARK_ITERATIONS)
            bs_b, _kgr_kuan_b_local, time_kuan_b = process_kuantisasi(kal_b, KUANTISASI_NUM_BITS, BENCHMARK_ITERATIONS)
            bs_ea, _kgr_kuan_ea_local, time_kuan_ea = process_kuantisasi(kal_ea, KUANTISASI_NUM_BITS, BENCHMARK_ITERATIONS)
            bs_eb, _kgr_kuan_eb_local, time_kuan_eb = process_kuantisasi(kal_eb, KUANTISASI_NUM_BITS, BENCHMARK_ITERATIONS)

            # KGR Kuantisasi kumulatif = panjang bitstream / (t_chanprob + t_kalman + t_kuantisasi)
            kgr_kuan_a = calculate_cumulative_kgr(len(bs_a), CHANPROB_TIME_SECONDS, time_kal_a, time_kuan_a)
            kgr_kuan_b = calculate_cumulative_kgr(len(bs_b), CHANPROB_TIME_SECONDS, time_kal_b, time_kuan_b)
            kgr_kuan_ea = calculate_cumulative_kgr(len(bs_ea), CHANPROB_TIME_SECONDS, time_kal_ea, time_kuan_ea)
            kgr_kuan_eb = calculate_cumulative_kgr(len(bs_eb), CHANPROB_TIME_SECONDS, time_kal_eb, time_kuan_eb)
            
            save_data_list(excel_kuan_dir, f"{v_name}_kuantisasi_alice.xlsx", [bs_a], "bitstream")
            save_data_list(excel_kuan_dir, f"{v_name}_kuantisasi_bob.xlsx", [bs_b], "bitstream")
            save_data_list(excel_kuan_dir, f"{v_name}_kuantisasi_evealice.xlsx", [bs_ea], "bitstream")
            save_data_list(excel_kuan_dir, f"{v_name}_kuantisasi_evebob.xlsx", [bs_eb], "bitstream")
            
            kdr_ab = calculate_kdr(bs_a, bs_b)
            kdr_eve = calculate_kdr(bs_ea, bs_eb)
            
            kuan_records.append({
                "skenario": skenario, "q": q, "r": r, "bb": bb,
                "kdr_ab": kdr_ab, "kdr_eve": kdr_eve,
                "time_alice": time_kuan_a, "time_bob": time_kuan_b, "time_evealice": time_kuan_ea, "time_evebob": time_kuan_eb,
                "kgr_alice": kgr_kuan_a, "kgr_bob": kgr_kuan_b, "kgr_evealice": kgr_kuan_ea, "kgr_evebob": kgr_kuan_eb,
                "total_bits_alice": len(bs_a), "total_bits_bob": len(bs_b),
                "total_bits_ea": len(bs_ea), "total_bits_eb": len(bs_eb)
            })
            
            # --- 4. BCH, Hash dan NIST Test (Simulasi modul) ---
            try:
                from bch_module import process_bch
                b_alice, b_bob, stats_ab = process_bch(bs_a, bs_b, apply_correction=True)
                b_ea, b_eb, stats_eve = process_bch(bs_ea, bs_eb, apply_correction=False)

                kgr_bch_ab = calculate_cumulative_kgr(
                    len(b_alice),
                    CHANPROB_TIME_SECONDS,
                    time_kal_a,
                    time_kuan_a,
                    stats_ab["time_bch"],
                )
                kgr_bch_eve = calculate_cumulative_kgr(
                    len(b_ea),
                    CHANPROB_TIME_SECONDS,
                    time_kal_ea,
                    time_kuan_ea,
                    stats_eve["time_bch"],
                )
                
                # Biar user bisa lihat list data array bits nya
                os.makedirs(os.path.join(skenario_out_dir, "data_excel_bch"), exist_ok=True)
                save_data_list(os.path.join(skenario_out_dir, "data_excel_bch"), f"{v_name}_bch_alice.xlsx", b_alice, "alice_bch_bits")
                save_data_list(os.path.join(skenario_out_dir, "data_excel_bch"), f"{v_name}_bch_bob.xlsx", b_bob, "bob_bch_bits")
                
                bch_records.append({
                    "skenario": skenario, "q": q, "r": r, "bb": bb,
                    "kdr_after_ab": stats_ab["kdr_after"], "kdr_after_eve": stats_eve["kdr_after"],
                    "kgr_bch_ab": kgr_bch_ab, "kgr_bch_eve": kgr_bch_eve,
                    "parity_bits_ab": stats_ab["parity_bits_sent"], "parity_bits_eve": stats_eve["parity_bits_sent"],
                    "total_bits_alice": stats_ab["total_bits_alice"], "total_bits_bob": stats_ab["total_bits_bob"],
                    "total_bits_ea": stats_eve["total_bits_alice"], "total_bits_eb": stats_eve["total_bits_bob"],
                    "error_bits_ab_before": stats_ab["error_bits_before"], "error_bits_eve_before": stats_eve["error_bits_before"],
                    "error_bits_ab_after": stats_ab["error_bits_after"], "error_bits_eve_after": stats_eve["error_bits_after"],
                    "corrected_bits_ab": stats_ab["corrected_bits"], "corrected_bits_eve": stats_eve["corrected_bits"],
                    "time_bch_ab": stats_ab["time_bch"], "time_bch_eve": stats_eve["time_bch"]
                })
                
                # --- 5. Hash & SHA & AES ---
                try: 
                    from hash_module import process_hash
                    # Proses Alice dan Bob
                    h_alice, h_bob, aes_ab, time_hash_ab, hash_metrics_ab = process_hash(b_alice, b_bob)
                    # Proses Eve-Alice dan Eve-Bob
                    h_ea, h_eb, aes_eve, time_hash_eve, hash_metrics_eve = process_hash(b_ea, b_eb)
                    
                    # Simpan Smua Keys (termasuk yg salah)
                    os.makedirs(os.path.join(skenario_out_dir, "data_excel_hash"), exist_ok=True)
                    save_data_list(os.path.join(skenario_out_dir, "data_excel_hash"), f"{v_name}_hash_alice.xlsx", h_alice, "AES_keys")
                    save_data_list(os.path.join(skenario_out_dir, "data_excel_hash"), f"{v_name}_hash_bob.xlsx", h_bob, "AES_keys")
                    save_data_list(os.path.join(skenario_out_dir, "data_excel_hash"), f"{v_name}_hash_evealice.xlsx", h_ea, "AES_keys")
                    save_data_list(os.path.join(skenario_out_dir, "data_excel_hash"), f"{v_name}_hash_evebob.xlsx", h_eb, "AES_keys")
                    
                    hash_records.append({
                        "skenario": skenario, "q": q, "r": r, "bb": bb,
                        "aes_count_ab": len(aes_ab), "aes_count_eve": len(aes_eve),
                        "keys_count_alice": hash_metrics_ab["keys_count_alice"], "keys_count_bob": hash_metrics_ab["keys_count_bob"],
                        "keys_count_ea": hash_metrics_eve["keys_count_alice"], "keys_count_eb": hash_metrics_eve["keys_count_bob"],
                        "total_key_bits_alice": hash_metrics_ab["total_key_bits_alice"],
                        "total_key_bits_bob": hash_metrics_ab["total_key_bits_bob"],
                        "total_key_bits_ea": hash_metrics_eve["total_key_bits_alice"],
                        "total_key_bits_eb": hash_metrics_eve["total_key_bits_bob"],
                        "matched_key_bits_ab": hash_metrics_ab["matched_key_bits"],
                        "matched_key_bits_eve": hash_metrics_eve["matched_key_bits"],
                        "final_key_alice": h_alice[0] if len(h_alice) > 0 else "N/A",
                        "final_key_bob": h_bob[0] if len(h_bob) > 0 else "N/A",
                        "final_key_ea": h_ea[0] if len(h_ea) > 0 else "N/A",
                        "final_key_eb": h_eb[0] if len(h_eb) > 0 else "N/A",
                        "time_hash_ab": time_hash_ab, "time_hash_eve": time_hash_eve
                    })
                    
                    # --- 6. Uji NIST ---
                    try:
                        from nist_module import process_nist
                        pass_ab, pval_ab, pass_rate_ab, pdist_ab, time_nist_ab = process_nist(aes_ab)
                        pass_eve, pval_eve, pass_rate_eve, pdist_eve, time_nist_eve = process_nist(aes_eve)

                        pdist_ab_str = ", ".join([f"{k}:{v}" for k, v in pdist_ab.items()])
                        pdist_eve_str = ", ".join([f"{k}:{v}" for k, v in pdist_eve.items()])
                        
                        nist_records.append({
                            "skenario": skenario, "q": q, "r": r, "bb": bb,
                            "passed_keys_ab": pass_ab, "passed_keys_eve": pass_eve,
                            "pval_ab": pval_ab, "pval_eve": pval_eve,
                            "pass_rate_ab": pass_rate_ab, "pass_rate_eve": pass_rate_eve,
                            "pval_dist_ab": pdist_ab_str, "pval_dist_eve": pdist_eve_str,
                            "time_nist_ab": time_nist_ab, "time_nist_eve": time_nist_eve
                        })
                    except Exception as e:
                        print("NIST Modul Error:", e)
                        
                except Exception as e:
                    print("Hash Modul Error:", e)
            except ImportError:
                print("Modul BCH Belum ditaruh di root folder!")
            
        # Mengukir Tabel Excel untuk Skenario 
        print(" == Menyusun File Table Laporan ==")
        build_kalman_excel(os.path.join(skenario_out_dir, "Laporan_Tabel_Kalman.xlsx"), kalman_records)
        build_kuantisasi_excel(os.path.join(skenario_out_dir, "Laporan_Tabel_Kuantisasi.xlsx"), kuan_records)
        
        if bch_records:
            build_bch_excel(os.path.join(skenario_out_dir, "Laporan_Tabel_BCH.xlsx"), bch_records)
        if hash_records:
            build_hash_excel(os.path.join(skenario_out_dir, "Laporan_Tabel_Hash.xlsx"), hash_records)
        if nist_records:
            build_nist_excel(os.path.join(skenario_out_dir, "Laporan_Tabel_NIST.xlsx"), nist_records)
            
        print("Selesai diproses untuk Skenario", skenario)
        
        # Tambahkan ke dictionary global untuk rekap semua skenario nanti
        global_kalman_records.extend(kalman_records)
        global_kuan_records.extend(kuan_records)
        if bch_records: global_bch_records.extend(bch_records)
        if hash_records: global_hash_records.extend(hash_records)
        if nist_records: global_nist_records.extend(nist_records)

    # === GENERATE GLOBAL REKAP ===
    print("\n==== MERANGKUM SELURUH SKENARIO KE DALAM SATU FILE EXCEL ====")
    rekap_excel_path = os.path.join(output_base, "Rekap_Evaluasi_SKG_Semua_Skenario.xlsx")
    
    # Buat workbook kosong
    rekap_wb = Workbook()
    
    # Hapus sheet default bawaan ('Sheet')
    if 'Sheet' in rekap_wb.sheetnames:
        rekap_wb.remove(rekap_wb['Sheet'])
        
    print("Menyusun sheet Kalman...")
    build_kalman_sheet(rekap_wb, global_kalman_records)
    
    print("Menyusun sheet Kuantisasi...")
    build_kuantisasi_sheet(rekap_wb, global_kuan_records)
    
    if global_bch_records:
        print("Menyusun sheet BCH...")
        build_bch_sheet(rekap_wb, global_bch_records)
        
    if global_hash_records:
        print("Menyusun sheet Hash...")
        build_hash_sheet(rekap_wb, global_hash_records)
        
    if global_nist_records:
        print("Menyusun sheet NIST...")
        build_nist_sheet(rekap_wb, global_nist_records)
        
    rekap_wb.save(rekap_excel_path)
    print(f"Selesai! File rekap global berhasil disimpan di: {rekap_excel_path}")

if __name__ == "__main__":
    main()
