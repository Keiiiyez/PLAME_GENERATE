import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

class PlameUniversalEngine:
    """Motor de reconocimiento basado en coincidencia física del Excel"""
    def __init__(self):
        self.ruc_empresa = "" 

    def analizar_fila(self, fila, tasas):
        # Convertimos la fila a texto para detectar palabras clave
        texto_fila = " ".join([str(c) for c in fila]).lower()
        
        # Extracción de montos reales redondeados
        nums_reales = []
        for c in fila:
            try:
                # Limpieza agresiva para capturar números financieros
                s = str(c).replace(',', '').replace('S/.', '').replace(' ', '').strip()
                res = re.search(r"(\d+\.\d+|\d+)", s)
                if res: nums_reales.append(round(float(res.group(1)), 2))
            except: pass

        # Identificación de DNI (8 dígitos) - Salta encabezados y totales
        dni = next((str(c).strip() for c in fila if re.match(r'^\d{8}$', str(c).strip())), None)
        if not dni or dni == self.ruc_empresa: 
            return None

        sueldo, onp, afp_ap = 0.0, 0.0, 0.0
        nums_sorted = sorted(list(set(nums_reales)), reverse=True)
        
        # --- Lógica de Retenciones (Cálculo Inverso) ---
        tasa_onp = 0.13
        tasa_afp_ap = tasas['aporte'] / 100

        for n in nums_sorted:
            # Evitamos años (2026) o montos muy pequeños
            if n < 300 or 2020 <= n <= 2030: continue 
            
            calc_onp = round(n * tasa_onp, 2)
            calc_afp = round(n * tasa_afp_ap, 2)
            
            # Si el 13% o el 10% calculado está presente en la fila, encontramos el básico
            if calc_onp in nums_reales:
                sueldo, onp = n, calc_onp
                break
            elif calc_afp in nums_reales:
                sueldo, afp_ap = n, calc_afp
                break
                
        # Si no hay retención clara, tomamos el monto más alto lógico (ej. para netos o básicos sin descuentos)
        if sueldo == 0.0:
            cands = [n for n in nums_sorted if 400 <= n <= 10000 and not (2020 <= n <= 2030)]
            sueldo = max(cands) if cands else 0.0

        # Seguro y Comisión dinámicos
        afp_sg, afp_cm = 0.0, 0.0
        if afp_ap > 0:
            tasa_sg = tasas['seguro'] / 100
            calc_sg = round(sueldo * tasa_sg, 2)
            if calc_sg in nums_reales:
                afp_sg = calc_sg
            
            for _, t_cm in tasas['comisiones'].items():
                calc_cm = round(sueldo * (t_cm/100), 2)
                if calc_cm in nums_reales and calc_cm > 0:
                    afp_cm = calc_cm
                    break

        # EsSalud con piso de ley (9%)
        ess_calc = max(101.70, round(sueldo * 0.09, 2))
        essalud = next((n for n in nums_reales if n == ess_calc or n == 101.70), ess_calc)

        horas = "168" # Valor estándar por defecto
        
        # 1. Caso Incapacidad Temporal o Subsidios
        if "inctemp" in texto_fila or "subsidio" in texto_fila:
            horas = "0"
        else:
            # 2. Análisis por celda para mayor precisión
            encontrado = False
            for celda in fila:
                val_str = str(celda).lower()
                
                match = re.search(r'(\d+)\s*h\b', val_str)
                if match:
                    horas = match.group(1)
                    encontrado = True
                    break
            
            # 3. Si no hay 'h' pero hay un número razonable después del sueldo
            if not encontrado:
                idx_basico = -1
                for i, v in enumerate(fila):
                    try:
                        if round(float(str(v).replace(',','')), 2) == sueldo:
                            idx_basico = i
                            break
                    except: continue
                
                # Buscamos en las 3 celdas siguientes al básico un número entre 1 y 240
                if idx_basico != -1:
                    for k in range(1, 4):
                        if idx_basico + k < len(fila):
                            c_val = str(fila[idx_basico + k])
                            m_h = re.search(r'(\d+)', c_val)
                            if m_h:
                                val_h = int(m_h.group(1))
                                if 1 <= val_h <= 240:
                                    horas = str(val_h)
                                    break

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
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file: return
        
        tasas = {
            'aporte': float(self.t_ap.get()), 'seguro': float(self.t_sg.get()),
            'comisiones': {'H': float(self.c_hab.get()), 'I': float(self.c_int.get()), 'P': float(self.c_pri.get()), 'F': float(self.c_pro.get())}
        }

        try:
            df = pd.read_excel(file, header=None)
            self.engine.ruc_empresa = self.ent_ruc.get()
            data_final = {}
            t_ess, t_onp, t_afp = 0, 0, 0

            self.tree.delete(*self.tree.get_children())
            for _, fila in df.iterrows():
                res = self.engine.analizar_fila(fila, tasas)
                if res and res['dni'] not in data_final:
                    data_final[res['dni']] = res
                    self.tree.insert("", "end", values=(res['dni'], f"{res['basico']:.2f}", f"{res['onp']:.2f}", f"{res['afp_ap']:.2f}", f"{res['afp_sg']:.2f}", f"{res['afp_cm']:.2f}", f"{res['essalud']:.2f}", res['horas']))
                    t_ess += res['essalud']
                    t_onp += res['onp']
                    t_afp += (res['afp_ap'] + res['afp_sg'] + res['afp_cm'])

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