import pandas as pd
excel_path = "Output/Rekap_Evaluasi_SKG_Semua_Skenario.xlsx"
try:
    xl = pd.ExcelFile(excel_path)
    print("Sheets:", xl.sheet_names)
    
    if "Hash Detail" in xl.sheet_names:
        df = pd.read_excel(excel_path, sheet_name="Hash Detail")
        print("\nHash Detail columns:", df.columns.tolist())
        print("Hash Detail head:\n", df.head(15).to_string())
    elif "Hash_SHA_AES" in xl.sheet_names:
        df = pd.read_excel(excel_path, sheet_name="Hash_SHA_AES")
        print("\nHash_SHA_AES columns:", df.columns.tolist())
        print("Hash_SHA_AES head:\n", df.head(15).to_string())
except Exception as e:
    print("Error:", e)
