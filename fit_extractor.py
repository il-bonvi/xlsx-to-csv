import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

HEADER_ROW = 116
UNIT_ROW = 117
DATA_START = 118


def convert_excel_to_csv():
    # Selezione Excel
    excel_path = filedialog.askopenfilename(
        title="Seleziona file Excel CPET",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not excel_path:
        return

    # Selezione output CSV
    csv_path = filedialog.asksaveasfilename(
        title="Salva CSV",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
        initialfile=os.path.splitext(os.path.basename(excel_path))[0] + "_pulito.csv"
    )

    if not csv_path:
        return

    try:
        df = pd.read_excel(excel_path, header=None)

        headers = df.iloc[HEADER_ROW].astype(str)
        units = df.iloc[UNIT_ROW].astype(str)

        merged_headers = [
            f"{h} ({u})" if u != "nan" else h
            for h, u in zip(headers, units)
        ]

        data = df.iloc[DATA_START:].copy()
        data.columns = merged_headers
        data = data.dropna(how="all")

        data.to_csv(csv_path, index=False)

        messagebox.showinfo(
            "Operazione completata",
            f"CSV creato correttamente:\n{csv_path}"
        )

    except Exception as e:
        messagebox.showerror("Errore", str(e))


# GUI base
root = tk.Tk()
root.title("CPET Excel → CSV")
root.geometry("420x180")
root.resizable(False, False)

label = tk.Label(
    root,
    text="Conversione CPET Excel → CSV\n(unisce header + unità)",
    font=("Arial", 11),
    pady=20
)
label.pack()

btn = tk.Button(
    root,
    text="Seleziona Excel e Converti",
    font=("Arial", 11),
    command=convert_excel_to_csv,
    width=30,
    height=2
)
btn.pack()

root.mainloop()