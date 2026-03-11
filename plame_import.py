import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

class PlameUniversalEngine:
    def __init__(self):
        self.ruc_empresa = "" 

    def analizar_fila(self, fila, tasas):
        # 1. Limpieza de datos extrema
        fila_str = []
        for c in fila:
            s = str(c).strip()
            # Si Pandas leyó la celda vacía como 'nan', la ignoramos
            if s.lower() == 'nan' or s == '': continue
            
            s = s.replace('S/.', '').replace(',', '')
            if s.endswith('.0'): s = s[:-2] # Quitar decimales de números enteros
            fila_str.append(s)
        
        texto_fila = " ".join(fila_str).lower()
        
        # FILTRO: Si la fila dice "sub total", "total fondo" o es cabecera, la saltamos
        if "sub total" in texto_fila or "total fondo" in texto_fila or "retenciones" in texto_fila:
            return None

        # 2. DETECCIÓN DE DNI (Corregido para ceros a la izquierda y notación limpia)
        dni = None
        for val in fila_str:
            if val.isdigit() and (len(val) == 8 or len(val) == 7):
                dni_candidato = val.zfill(8)
                if dni_candidato != self.ruc_empresa:
                    dni = dni_candidato
                    break

        if not dni: return None

        # 3. EXTRACCIÓN DE MONTOS (Sueldos y Retenciones)
        nums_reales = []
        for s in fila_str:
            # Verificamos matemáticamente si el texto es un número válido (ej. "1130.00")
            if s.replace('.', '', 1).isdigit():
                val = round(float(s), 2)
                # Filtrar para no confundir con DNI o Años
                if val < 20000 and not (2020 <= val <= 2030) and val != float(dni):
                    nums_reales.append(val)

        if not nums_reales: return None

        sueldo, onp, afp_ap = 0.0, 0.0, 0.0
        nums_sorted = sorted(list(set(nums_reales)), reverse=True)
        
        tasa_onp = 0.13
        tasa_afp_ap = tasas['aporte'] / 100

        # Buscar coincidencia matemática Sueldo <-> Retención
        for n in nums_sorted:
            if n < 50: continue # Montos muy pequeños no son sueldos base
            
            c_onp = round(n * tasa_onp, 2)
            c_afp = round(n * tasa_afp_ap, 2)
            
            if any(abs(x - c_onp) <= 0.10 for x in nums_reales):
                sueldo = n
                onp = next(x for x in nums_reales if abs(x - c_onp) <= 0.10)
                break
            elif any(abs(x - c_afp) <= 0.10 for x in nums_reales):
                sueldo = n
                afp_ap = next(x for x in nums_reales if abs(x - c_afp) <= 0.10)
                break

        # Caso especial: Si no hay retención (como la fila de liquidación)
        if sueldo == 0.0 and nums_sorted:
            sueldo = nums_sorted[0]

        # 4. SEGURO Y COMISIÓN (AFP)
        afp_sg, afp_cm = 0.0, 0.0
        if afp_ap > 0:
            tasa_sg = tasas['seguro'] / 100
            for n in nums_reales:
                if abs(n - (sueldo * tasa_sg)) <= 0.05:
                    afp_sg = n
                    break
            for _, t_cm in tasas['comisiones'].items():
                calc_cm = sueldo * (t_cm/100)
                found_cm = [x for x in nums_reales if abs(x - calc_cm) <= 0.05]
                if found_cm:
                    afp_cm = found_cm[0]
                    break

        # 5. ESSALUD Y HORAS
        ess_calc = round(sueldo * 0.09, 2)
        # Si encuentra 101.70 exacto en la fila, lo usa
        essalud = 101.70 if any(abs(x - 101.70) < 0.01 for x in nums_reales) else max(101.70, ess_calc)
        if sueldo == 0: essalud = 0.0

        horas = "168"
        if "inctemp" in texto_fila or "0 d, 0 h" in texto_fila:
            horas = "0"
        else:
            match_h = re.search(r'(\d+)\s*h', texto_fila)
            if match_h: horas = match_h.group(1)

        return {
            'dni': dni, 'basico': sueldo, 'onp': onp,
            'afp_ap': afp_ap, 'afp_sg': afp_sg, 'afp_cm': afp_cm,
            'essalud': essalud, 'horas': horas
        }

class AppV13(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IMPORTADOR PLAMES")
        self.geometry("1150x700")
        self.configure(bg="#f1f5f9")
        self.engine = PlameUniversalEngine()
        self.init_ui()

    def init_ui(self):
        # Panel Lateral de Tasas
        side = tk.Frame(self, bg="white", width=220, relief="ridge", borderwidth=1)
        side.pack(side="left", fill="y", padx=10, pady=10)
        
        tk.Label(side, text="TASAS AFP MES", font=("Arial", 10, "bold"), bg="white").pack(pady=15)
        self.t_ap = self.add_rate(side, "Aporte %", "10.00")
        self.t_sg = self.add_rate(side, "Seguro %", "1.37")
        tk.Label(side, text="Comisiones flujo:", font=("Arial", 8, "italic"), bg="white").pack(pady=5)
        self.c_hab = self.add_rate(side, "Habitat", "1.47")
        self.c_int = self.add_rate(side, "Integra", "1.55")
        self.c_pri = self.add_rate(side, "Prima", "1.60")
        self.c_pro = self.add_rate(side, "Profuturo", "1.69")

        # Cuerpo Principal
        main = tk.Frame(self, bg="#f1f5f9")
        main.pack(side="right", expand=True, fill="both")

        top = tk.Frame(main, bg="#f1f5f9", pady=15)
        top.pack()
        self.ent_ruc = self.add_input(top, "RUC:", "", 0)
        self.ent_per = self.add_input(top, "Periodo (AAAAMM):", "202601", 1)

        tk.Button(main, text="PROCESAR EXCEL Y GENERAR TXT", command=self.run, bg="#2563eb", fg="white", font=("Arial", 10, "bold"), pady=10).pack(fill="x", padx=40)

        # Tabla de Resultados
        self.tree = ttk.Treeview(main, columns=("DNI", "Sueldo", "ONP", "Aporte", "Seguro", "Comisión", "EsSalud", "Horas"), show='headings')
        for col in self.tree["columns"]: self.tree.heading(col, text=col); self.tree.column(col, width=100, anchor="center")
        self.tree.pack(expand=True, fill="both", padx=20, pady=15)

        # Totales para cuadre
        self.lbl_tot = tk.Label(main, text="TOTALES CARGADOS -> EsSalud: S/. 0.00 | ONP: S/. 0.00 | AFP: S/. 0.00", bg="#cbd5e1", font=("Arial", 10, "bold"), pady=10)
        self.lbl_tot.pack(fill="x")

    def add_rate(self, master, txt, val):
        f = tk.Frame(master, bg="white"); f.pack(pady=2)
        tk.Label(f, text=txt, width=12, anchor="w", bg="white").pack(side="left", padx=5)
        e = tk.Entry(f, width=8, justify="center"); e.insert(0, val); e.pack(side="right", padx=5); return e

    def add_input(self, master, txt, val, c):
        tk.Label(master, text=txt, bg="#f1f5f9", font=("Arial", 9, "bold")).grid(row=0, column=c*2, padx=5)
        e = tk.Entry(master, justify="center"); e.insert(0, val); e.grid(row=0, column=c*2+1, padx=10); return e

    def run(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file: return
        
        tasas = {
            'aporte': float(self.t_ap.get()), 'seguro': float(self.t_sg.get()),
            'comisiones': {'H': float(self.c_hab.get()), 'I': float(self.c_int.get()), 'P': float(self.c_pri.get()), 'F': float(self.c_pro.get())}
        }

        try:
            # CORRECCIÓN VITAL: Leer todo como string (dtype=str) para que no rompa los DNI
            df = pd.read_excel(file, header=None, dtype=str)
            self.engine.ruc_empresa = self.ent_ruc.get()
            data_final = {}
            t_ess, t_onp, t_afp = 0, 0, 0

            self.tree.delete(*self.tree.get_children())
            for _, fila in df.iterrows():
                # Convertimos la fila de Pandas a lista normal
                res = self.engine.analizar_fila(fila.tolist(), tasas)
                
                if res and res['dni'] not in data_final:
                    data_final[res['dni']] = res
                    self.tree.insert("", "end", values=(res['dni'], f"{res['basico']:.2f}", f"{res['onp']:.2f}", f"{res['afp_ap']:.2f}", f"{res['afp_sg']:.2f}", f"{res['afp_cm']:.2f}", f"{res['essalud']:.2f}", res['horas']))
                    t_ess += res['essalud']
                    t_onp += res['onp']
                    t_afp += (res['afp_ap'] + res['afp_sg'] + res['afp_cm'])

            if not data_final:
                messagebox.showwarning("Aviso", "No se detectaron trabajadores. Revisa el Excel.")
                return

            self.lbl_tot.config(text=f"TOTALES CARGADOS -> EsSalud: S/. {t_ess:.2f} | ONP: S/. {t_onp:.2f} | AFP: S/. {t_afp:.2f}")
            self.save_txt(data_final)
            messagebox.showinfo("Éxito", f"Se procesaron {len(data_final)} empleados y se generaron los TXT.")
            os.startfile(os.getcwd())
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al procesar: {e}")

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
    AppV13().mainloop()