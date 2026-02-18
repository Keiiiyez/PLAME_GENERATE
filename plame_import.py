import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import glob


COMISIONES_AFP = {
    'INTEGRA':   {'comision': 0.0155, 'seguro': 0.0137, 'aporte': 0.10},
    'PRIMA':     {'comision': 0.0000, 'seguro': 0.0137, 'aporte': 0.10},
    'PROFUTURO': {'comision': 0.0000, 'seguro': 0.0137, 'aporte': 0.10},
    'HABITAT':   {'comision': 0.0000, 'seguro': 0.0137, 'aporte': 0.10},
}
RMV = 1025.00
TOPE_SEGURO_AFP = 12209.11

class PlameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador PLAME SUNAT")
        self.root.geometry("500x520")
        self.root.configure(bg="#f0f3f5")
        
        
        tk.Label(root, text="GENERADOR PLAME", 
                 font=("Arial", 14, "bold"), bg="#f0f3f5", fg="#2c3e50").pack(pady=15)
        
        
        tk.Label(root, text="RUC Empresa:", bg="#f0f3f5").pack()
        self.ent_ruc = tk.Entry(root, justify='center', width=20, font=("Arial", 10))
        self.ent_ruc.insert(0, "")
        self.ent_ruc.pack(pady=5)
        
        tk.Label(root, text="Periodo (AAAAMM):", bg="#f0f3f5").pack()
        self.ent_periodo = tk.Entry(root, justify='center', width=20, font=("Arial", 10))
        self.ent_periodo.insert(0, "202601")
        self.ent_periodo.pack(pady=5)
        
       
        self.btn_gen = tk.Button(root, text="PROCESAR PLANILLA EXCEL", 
                                command=self.procesar, bg="#27ae60", fg="white", 
                                font=("Arial", 10, "bold"), height=2, width=35)
        self.btn_gen.pack(pady=20)

        self.btn_folder = tk.Button(root, text="ABRIR CARPETA DE RESULTADOS", 
                                   command=self.abrir_carpeta, bg="#3498db", fg="white", 
                                   font=("Arial", 9), width=30)
        self.btn_folder.pack(pady=5)
        
        self.lbl_status = tk.Label(root, text="Listo para procesar", bg="#f0f3f5", fg="gray")
        self.lbl_status.pack(pady=10)

    def abrir_carpeta(self):
        os.startfile(os.getcwd())

    def eliminar_antiguos(self):
        formatos = ['*.rem', '*.tra', '*.jor', '*.snl']
        for f in formatos:
            for archivo in glob.glob(f):
                try: os.remove(archivo)
                except: pass

    def limpiar_monto(self, valor):
        try:
            val = str(valor).replace(',', '').strip()
            return float(val) if val not in ['-', '', 'None', 'nan', '#¡REF!'] else 0.0
        except: return 0.0

    def procesar(self):
        ruc = self.ent_ruc.get().strip()
        periodo = self.ent_periodo.get().strip()
        
        if len(ruc) != 11 or len(periodo) != 6:
            messagebox.showwarning("Error", "RUC o Periodo con formato incorrecto.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path: return

        try:
            self.eliminar_antiguos()
            df_raw = pd.read_excel(file_path, header=None)
            resultados = []
            
            for i, fila in df_raw.iterrows():
                dni = None
                idx_dni = -1
                for idx, celda in enumerate(fila):
                    if re.match(r'^\d{8}$', str(celda).strip()):
                        dni = str(celda).strip()
                        idx_dni = idx
                        break
                
                if not dni: continue 

               
                basico = self.limpiar_monto(fila[idx_dni + 7])
                
                asig_fam = self.limpiar_monto(fila[idx_dni + 13]) if (idx_dni + 13) < len(fila) else 0
                total_rem = basico + asig_fam
                
                
                celda_horas = str(fila[idx_dni + 8])
                es_subsidio = "INC" in celda_horas.upper() or "SUB" in celda_horas.upper()
                
                
                horas = "168"
                match_h = re.search(r'(\d+)\s*h', celda_horas)
                if match_h: horas = match_h.group(1)
                elif es_subsidio: horas = "0"

               
                info_pension = str(fila[idx_dni + 17]).upper()
                tributos = {}
                
                if any(x in info_pension for x in ["INTEGRA", "PRIMA", "PROFUTURO", "HABITAT"]):
                    afp_nom = next(a for a in COMISIONES_AFP if a in info_pension)
                    c = COMISIONES_AFP[afp_nom]
                    tributos["0608"] = round(total_rem * 0.10, 2)
                    tributos["0601"] = round(min(total_rem, TOPE_SEGURO_AFP) * c['seguro'], 2)
                    if c['comision'] > 0:
                        tributos["0606"] = round(total_rem * c['comision'], 2)
                else:
                    monto_onp = self.limpiar_monto(fila[idx_dni + 16])
                    if monto_onp > 0: tributos["0607"] = round(monto_onp, 2)

                tributos["0804"] = round(max(total_rem * 0.09, RMV * 0.09), 2)

                resultados.append({
                    'DNI': dni, 'BASICO': basico, 'ASIG_FAM': asig_fam,
                    'TRIBUTOS': tributos, 'HORAS': horas, 'SUBSIDIO': es_subsidio
                })

            if resultados:
                base = f"0601{periodo}{ruc}"
                # Guardar REM
                with open(f"{base}.rem", "w", encoding="ansi") as f:
                    for r in resultados:
                        f.write(f"01|{r['DNI']}|0121|{r['BASICO']:.2f}|{r['BASICO']:.2f}|\r\n")
                        if r['ASIG_FAM'] > 0:
                            f.write(f"01|{r['DNI']}|0201|{r['ASIG_FAM']:.2f}|{r['ASIG_FAM']:.2f}|\r\n")
                
                # Guardar TRA
                with open(f"{base}.tra", "w", encoding="ansi") as f:
                    for r in resultados:
                        for cod, mon in r['TRIBUTOS'].items():
                            f.write(f"01|{r['DNI']}|{cod}|{mon:.2f}|{mon:.2f}|\r\n")
                
                # Guardar JOR
                with open(f"{base}.jor", "w", encoding="ansi") as f:
                    for r in resultados:
                        f.write(f"01|{r['DNI']}|{r['HORAS']}|0|0|0|\r\n")
                
                # Guardar SNL (Subsidios)
                subs = [r for r in resultados if r['SUBSIDIO']]
                if subs:
                    with open(f"{base}.snl", "w", encoding="ansi") as f:
                        for r in subs:
                            
                            f.write(f"01|{r['DNI']}|21|30|\r\n")

                messagebox.showinfo("Éxito", f"Se generaron archivos para {len(resultados)} trabajadores.")
                self.lbl_status.config(text="Proceso completado.", fg="green")
            else:
                messagebox.showwarning("Aviso", "No se detectó información procesable.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un problema: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlameApp(root)
    root.mainloop()