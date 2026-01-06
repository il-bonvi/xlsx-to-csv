import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def normalize_decimals(df):
    """
    Converte numeri scritti come stringhe con virgola decimale
    (es. '12,34') in float 12.34
    """
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = (
                df[col]
                .str.replace(".", "", regex=False)   # rimuove separatori migliaia
                .str.replace(",", ".", regex=False)  # converte decimali
            )
            df[col] = pd.to_numeric(df[col], errors="ignore")
    return df


def process_xlsx_to_csv():
    xlsx_path = filedialog.askopenfilename(
        title="Seleziona file Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not xlsx_path:
        return

    output_dir = filedialog.askdirectory(
        title="Seleziona cartella di esportazione CSV"
    )

    if not output_dir:
        return

    try:
        xls = pd.ExcelFile(xlsx_path)
        base_name = os.path.splitext(os.path.basename(xlsx_path))[0]

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = normalize_decimals(df)

            safe_sheet = sheet.replace(" ", "_")
            output_path = os.path.join(
                output_dir, f"{base_name}_{safe_sheet}.csv"
            )

            df.to_csv(
                output_path,
                index=False,
                sep=",",
                decimal=".",
                encoding="utf-8"
            )

        messagebox.showinfo(
            "OK",
            f"Creati {len(xls.sheet_names)} file CSV (formato US)"
        )

    except Exception as e:
        messagebox.showerror("Errore", str(e))


# GUI
root = tk.Tk()
root.title("Excel → CSV (US format)")
root.geometry("420x180")
root.resizable(False, False)

label = tk.Label(
    root,
    text="Apre XLSX ed esporta CSV formato US\nseparatore ,  decimali .",
    font=("Arial", 11),
    pady=20
)
label.pack()

btn = tk.Button(
    root,
    text="Apri Excel e crea CSV",
    font=("Arial", 11),
    command=process_xlsx_to_csv,
    width=30,
    height=2
)
btn.pack()

root.mainloop()