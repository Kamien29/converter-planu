

import pandas as pd
import re
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime

# dni tygodnia
WEEKDAYS = ["poniedziałek", "wtorek", "środa", "czwartek", "piątek"]

# regex do nagłówka klasy (np. "1AT - 1.09.2025")
HEADER_CLASS_RE = re.compile(r"^\s*([^\s–-]+)\s*[-–]\s*\d{1,2}\.\d{1,2}\.\d{2,4}")

# regex do godzin (np. "7:10 - 7:55")
TIME_RE = re.compile(r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})")


def normalize_time(t: str):
    """Normalizacja godziny do formatu HH:MM:00"""
    if not t:
        return None
    m = re.match(r"^(\d{1,2}):(\d{2})$", t.strip())
    if not m:
        return None
    return f"{int(m.group(1)):02d}:{m.group(2)}:00"


def escape_sql(s):
    """Przygotuj string do SQL (podwójne apostrofy, bez nowych linii)"""
    if s is None:
        return ""
    s = str(s).replace("\n", " ").replace("\r", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s.replace("'", "''")


def find_weekday_columns(row_values):
    """
    Wyszukuje kolumny z dniami tygodnia
    row_values: lista stringów z jednego wiersza
    return: dict {col_index: weekday_name}
    """
    mapping = {}
    for idx, cell in enumerate(row_values):
        if not cell:
            continue
        cell_clean = str(cell).strip().lower()
        for wd in WEEKDAYS:
            if wd == cell_clean or cell_clean.startswith(wd[:3]):
                mapping[idx] = wd.capitalize()
                break
    return mapping if len(mapping) >= 3 else {}


def parse_excel_to_plan(path, log_append):
    """
    Wczytuje wszystkie arkusze Excela i zwraca listę rekordów planu
    plan = [{klasa, dzien, start, end, przedmiot, sheet, row, col}, ...]
    """
    plan = []
    errors = []

    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        errors.append(f"Błąd otwarcia pliku: {e}")
        return plan, errors

    for sheetname in xls.sheet_names:
        try:
            df = xls.parse(sheetname, header=None, dtype=object)
        except Exception as e:
            errors.append(f"Nie można przeczytać arkusza '{sheetname}': {e}")
            continue

        current_class = None
        weekday_cols = {}

        for i in range(len(df)):
            row = df.iloc[i].fillna("").astype(str).tolist()

            # 1) Nagłówek klasy
            for col_idx, cell in enumerate(row):
                m = HEADER_CLASS_RE.match(cell)
                if m:
                    current_class = m.group(1)
                    log_append(f"[{sheetname}] Znaleziono klasę '{current_class}' w wierszu {i+1}")
                    weekday_cols = {}
                    break

            # 2) Wiersz z dniami tygodnia
            if not weekday_cols:
                cand = find_weekday_columns(row)
                if cand:
                    weekday_cols = cand
                    log_append(f"[{sheetname}] Dni tygodnia w wierszu {i+1}: {weekday_cols}")
                    continue

            # 3) Wiersze z godzinami
            time_match = None
            time_col_idx = None
            for col_idx, cell in enumerate(row):
                m = TIME_RE.search(cell)
                if m:
                    time_match = m
                    time_col_idx = col_idx
                    break

            if time_match:
                if not current_class:
                    current_class = "UNKNOWN"
                    errors.append(f"[{sheetname}] Wiersz {i+1}: godziny znalezione, ale brak klasy → ustawiono UNKNOWN")

                if not weekday_cols:
                    errors.append(f"[{sheetname}] Wiersz {i+1}: godziny znalezione, ale brak dni tygodnia")
                    continue

                start = normalize_time(time_match.group(1))
                end = normalize_time(time_match.group(2))
                if not start or not end:
                    errors.append(f"[{sheetname}] Wiersz {i+1}: zły format godziny '{time_match.group(0)}'")
                    continue

                # dodajemy przedmioty
                for col_idx, weekday in weekday_cols.items():
                    if col_idx >= len(row):
                        continue
                    subj = row[col_idx].strip()
                    if subj:
                        plan.append({
                            "klasa": current_class,
                            "dzien": weekday,
                            "start": start,
                            "end": end,
                            "przedmiot": escape_sql(subj),
                            "sheet": sheetname,
                            "row": i+1,
                            "col": col_idx+1
                        })

    return plan, errors


def write_sql(plan, out_path):
    """Zapis planu do pliku SQL"""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("-- Wygenerowano: " + datetime.now().isoformat() + "\n")
        f.write("CREATE TABLE IF NOT EXISTS plan (\n")
        f.write("  id INT AUTO_INCREMENT PRIMARY KEY,\n")
        f.write("  klasa VARCHAR(50),\n")
        f.write("  dzien VARCHAR(20),\n")
        f.write("  godzina_start TIME,\n")
        f.write("  godzina_koniec TIME,\n")
        f.write("  przedmiot VARCHAR(500),\n")
        f.write("  sheet VARCHAR(100),\n")
        f.write("  sheet_row INT,\n")
        f.write("  sheet_col INT\n")
        f.write(");\n\n")

        for i, r in enumerate(plan, start=1):
            f.write(
                f"INSERT INTO plan (id, klasa, dzien, godzina_start, godzina_koniec, przedmiot, sheet, sheet_row, sheet_col) "
                f"VALUES ({i}, '{escape_sql(r['klasa'])}', '{escape_sql(r['dzien'])}', '{r['start']}', '{r['end']}', "
                f"'{r['przedmiot']}', '{escape_sql(r['sheet'])}', {r['row']}, {r['col']});\n"
            )


# ---------------- GUI ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Konwerter planu → SQL")

        frm = tk.Frame(root)
        frm.pack(padx=10, pady=8, fill=tk.X)

        tk.Label(frm, text="Plik Excel:").grid(row=0, column=0, sticky="w")
        self.entry_file = tk.Entry(frm, width=60)
        self.entry_file.grid(row=0, column=1, padx=5)
        tk.Button(frm, text="Wybierz...", command=self.choose_file).grid(row=0, column=2)

        tk.Label(frm, text="Plik SQL (wyjście):").grid(row=1, column=0, sticky="w")
        self.entry_sql = tk.Entry(frm, width=60)
        self.entry_sql.grid(row=1, column=1, padx=5)
        tk.Button(frm, text="Wybierz...", command=self.choose_out).grid(row=1, column=2)

        self.btn_run = tk.Button(root, text="Konwertuj", command=self.run)
        self.btn_run.pack(pady=6)

        self.log = scrolledtext.ScrolledText(root, width=90, height=20, wrap=tk.WORD)
        self.log.pack(padx=10, pady=8)

    def choose_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, p)
            out_default = str(Path(p).with_suffix(".sql"))
            self.entry_sql.delete(0, tk.END)
            self.entry_sql.insert(0, out_default)

    def choose_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".sql", filetypes=[("SQL", "*.sql")])
        if p:
            self.entry_sql.delete(0, tk.END)
            self.entry_sql.insert(0, p)

    def append_log(self, s):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.insert(tk.END, f"[{ts}] {s}\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    def run(self):
        in_path = self.entry_file.get().strip()
        out_path = self.entry_sql.get().strip()
        if not in_path:
            messagebox.showerror("Błąd", "Wybierz plik Excel.")
            return
        if not out_path:
            messagebox.showerror("Błąd", "Wybierz plik wyjściowy SQL.")
            return

        self.btn_run.config(state=tk.DISABLED)
        self.log.delete(1.0, tk.END)
        self.append_log(f"Start konwersji: {in_path}")

        try:
            plan, errors = parse_excel_to_plan(in_path, self.append_log)
            self.append_log(f"Znaleziono {len(plan)} wpisów.")

            if plan:
                write_sql(plan, out_path)
                self.append_log(f"Zapisano SQL do: {out_path}")
            else:
                self.append_log("Brak wpisów do zapisania (plan pusty).")

            if errors:
                self.append_log("=== UWAGI ===")
                for e in errors:
                    self.append_log("! " + e)
                messagebox.showwarning("Zakończono z uwagami", f"Zapisano {len(plan)} rekordów, {len(errors)} uwag.")
            else:
                messagebox.showinfo("Sukces", f"Zapisano {len(plan)} rekordów.")
        except Exception as e:
            tb = traceback.format_exc()
            self.append_log("Błąd krytyczny: " + str(e))
            self.append_log(tb)
            messagebox.showerror("Błąd krytyczny", f"{e}\nSzczegóły w logu.")
        finally:
            self.btn_run.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.geometry("950x620")
    root.mainloop()
