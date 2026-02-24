import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

class PlameUniversalEngine:
    """Generador"""
    def __init__(self):
        self.ruc_empresa = "" 
    def analizar_fila(self, fila):
        texto_fila = " ".join([str(c) for c in fila]).lower()
        
        # Extracción segura conservando el orden de izquierda a derecha
        nums_originales = []
        for c in fila:
            if isinstance(c, (int, float)) and pd.notna(c):
                nums_originales.append(round(float(c), 2))
            else:
                try:
                    s = str(c).replace(',', '').replace('S/.', '').strip()
                    res = re.search(r"(\d+\.\d+|\d+)", s)
                    if res: nums_originales.append(round(float(res.group(1)), 2))
                except: pass

        # DNI: Buscamos 8 dígitos seguidos
        dni = next((str(c).strip() for c in fila if re.match(r'^\d{8}$', str(c).strip())), None)
        if not dni or dni == self.ruc_empresa: 
            return None

        # --- LÓGICA MATEMÁTICA UNIVERSAL ---
        sueldo = 0.0
        onp = 0.0
        afp_ap = 0.0
        
        # Ordenamos de mayor a menor solo para buscar el Sueldo Base
        nums_sorted = sorted(list(set(nums_originales)), reverse=True)
        
        for n in nums_sorted:
            if n < 300: continue # Ignoramos montos muy bajos como sueldo base
            
            calc_onp = round(n * 0.13, 2)
            calc_afp = round(n * 0.10, 2)
            
            # Si el 13% de este número está en la fila, es ONP
            if calc_onp in nums_sorted:
                sueldo = n
                onp = calc_onp
                break
            # Si el 10% de este número está en la fila, es AFP
            elif calc_afp in nums_sorted:
                sueldo = n
                afp_ap = calc_afp
                break
                
        # Si no hubo retención (Ej: practicante), tomamos el monto más alto lógico
        if sueldo == 0.0:
            cands = [n for n in nums_sorted if 400 <= n <= 10000 and not (1940 <= n <= 2030)]
            sueldo = max(cands) if cands else 0.0

        # Seguro y Comisión de AFP
        afp_sg = 0.0
        afp_cm = 0.0
        if afp_ap > 0:
            # Recorremos de izquierda a derecha para asegurar que Seguro va antes que Comisión
            for n in nums_originales:
                if round(sueldo * 0.005, 2) <= n <= round(sueldo * 0.025, 2):
                    if afp_sg == 0.0:
                        afp_sg = n
                    elif afp_cm == 0.0 and n != afp_sg:
                        afp_cm = n

        # EsSalud (Siempre el 9%, con un piso de 101.70 por ley)
        essalud_calc = max(101.70, round(sueldo * 0.09, 2))
        essalud = essalud_calc if essalud_calc in nums_sorted else 101.70

        # Horas (Extrae el número antes de la "h" o "d, x h")
        horas = "168"
        match = re.search(r'(?:^|\s)(\d+)\s*h\b', texto_fila)
        if match:
            horas = match.group(1)
        elif "inctemp" in texto_fila or "subsidio" in texto_fila:
            horas = "0"

        return {
            'dni': dni,
            'nom': str(fila[2])[:30] if pd.notna(fila[2]) else "TRABAJADOR",
            'basico': sueldo,
            'onp': onp,
            'afp_ap': afp_ap,
            'afp_sg': afp_sg,
            'afp_cm': afp_cm,
            'essalud': essalud,
            'horas': horas
        }

class AppV13:
    def __init__(self, root):
        self.root = root
        self.root.title("PLAME Generador")
        self.root.geometry("1050x650")
        self.root.configure(bg="#f8fafc")
        self.engine = PlameUniversalEngine()
        
        
        header = tk.Frame(root, bg="#0f172a", height=80)
        header.pack(fill="x")
        tk.Label(header, text="CONVERTIDOR PLAME", font=("Segoe UI", 16, "bold"), fg="white", bg="#0f172a").pack(pady=20)

        
        frame_input = tk.Frame(root, bg="#f8fafc", pady=20)
        frame_input.pack()
        self.ent_ruc = self.add_input(frame_input, "RUC:", self.engine.ruc_empresa, 0)
        self.ent_per = self.add_input(frame_input, "Periodo:", "202601", 1)

        tk.Button(root, text="PROCESAR Y GENERAR", command=self.run, bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"), padx=25, pady=10, relief="flat").pack()

        
        style = ttk.Style(); style.theme_use("clam")
        self.tree = ttk.Treeview(root, columns=("DNI", "Sueldo", "ONP", "AFP", "Seguro", "Comisión", "EsSalud", "Horas"), show='headings')
        for col in self.tree["columns"]: 
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=20, pady=20)

    def add_input(self, master, txt, dft, c):
        tk.Label(master, text=txt, bg="#f8fafc", font=("Segoe UI", 9, "bold")).grid(row=0, column=c*2, padx=5)
        e = tk.Entry(master, font=("Segoe UI", 10), justify="center"); e.insert(0, dft); e.grid(row=0, column=c*2+1, padx=15)
        return e

    def run(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file: return
        
        try:
            df = pd.read_excel(file, header=None)
            self.engine.ruc_empresa = self.ent_ruc.get()
            data_final = {}

            for _, fila in df.iterrows():
                res = self.engine.analizar_fila(fila)
                # omitir duplicados de liquidación al final del Excel
                if not res or res['dni'] in data_final: continue
                data_final[res['dni']] = res

            self.save_txt(data_final)
            
            self.tree.delete(*self.tree.get_children())
            for d, v in data_final.items():
                self.tree.insert("", "end", values=(d, f"{v['basico']:.2f}", f"{v['onp']:.2f}", f"{v['afp_ap']:.2f}", f"{v['afp_sg']:.2f}", f"{v['afp_cm']:.2f}", f"{v['essalud']:.2f}", v['horas']))
            
            messagebox.showinfo("Proceso Exitoso", f"Se procesaron {len(data_final)} trabajadores.")
            os.startfile(os.getcwd())
            
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al procesar el documento: {e}")

    def save_txt(self, data):
        base = f"0601{self.ent_per.get()}{self.ent_ruc.get()}"
        with open(base+".rem", "w") as f:
            for d, v in data.items(): f.write(f"01|{d}|0121|{v['basico']:.2f}|{v['basico']:.2f}|\r\n")
        with open(base+".tra", "w") as f:
            for d, v in data.items():
                if v['onp'] > 0: f.write(f"01|{d}|0607|{v['onp']:.2f}|{v['onp']:.2f}|\r\n")
                if v['afp_ap'] > 0: f.write(f"01|{d}|0608|{v['afp_ap']:.2f}|{v['afp_ap']:.2f}|\r\n")
                if v['afp_sg'] > 0: f.write(f"01|{d}|0601|{v['afp_sg']:.2f}|{v['afp_sg']:.2f}|\r\n")
                if v['afp_cm'] > 0: f.write(f"01|{d}|0606|{v['afp_cm']:.2f}|{v['afp_cm']:.2f}|\r\n")
                f.write(f"01|{d}|0804|{v['essalud']:.2f}|{v['essalud']:.2f}|\r\n")
        with open(base+".jor", "w") as f:
            for d, v in data.items(): f.write(f"01|{d}|{v['horas']}|0|0|0|\r\n")

if __name__ == "__main__":
    root = tk.Tk(); AppV13(root); root.mainloop()