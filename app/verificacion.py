import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
from collections import Counter

def run_comparator_gui():
    C_BG      = "#0F172A"
    C_SURFACE = "#1E293B"
    C_BORDER  = "#334155"
    C_ACCENT  = "#0EA5E9"
    C_SUCCESS = "#10B981"
    C_WARNING = "#F59E0B"
    C_ERROR   = "#EF4444"
    C_TEXT    = "#F1F5F9"
    C_MUTED   = "#64748B"
    C_INPUT   = "#0B1120"
    F         = "Segoe UI"

    root = tk.Tk()
    root.title("Validador y Comparador Profesional - SUNAT Pro")
    root.geometry("950x820")
    root.configure(bg=C_BG)
    root.minsize(800, 650)

    # Estilo de pestañas
    style = ttk.Style()
    style.theme_use('default')
    style.configure('TNotebook', background=C_BG, borderwidth=0)
    style.configure('TNotebook.Tab', background=C_SURFACE, foreground=C_TEXT, padding=[16, 8], font=(F, 9, 'bold'))
    style.map('TNotebook.Tab', background=[('selected', C_ACCENT)], foreground=[('selected', '#0F172A')])

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

    # ==================== PESTAÑA 1: COMPARAR ARCHIVOS EXCEL ====================
    tab1 = tk.Frame(notebook, bg=C_BG)
    notebook.add(tab1, text="  📁 Comparar Archivos Excel (Detallado)  ")

    pad1 = tk.Frame(tab1, bg=C_BG)
    pad1.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

    tk.Label(pad1, text="Auditoría y Auditoría de Datos Incorrectos", font=(F, 18, "bold"), bg=C_BG, fg=C_TEXT).pack(anchor=tk.W)
    tk.Label(pad1, text="Compara dos archivos Excel y muestra exactamente qué campos o datos específicos cambiaron o son incorrectos.", font=(F, 9), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(2, 12))
    tk.Frame(pad1, height=1, bg=C_BORDER).pack(fill=tk.X, pady=(0, 15))

    # --- ARCHIVO 1 (Versión 1) ---
    tk.Label(pad1, text="EXCEL VERSIÓN 1 (Original / Base)", font=(F, 8, "bold"), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(0, 4))
    f1_var = tk.StringVar()
    wrap1 = tk.Frame(pad1, bg=C_BG)
    wrap1.pack(fill=tk.X, pady=(0, 10))
    
    tk.Entry(wrap1, textvariable=f1_var, font=(F, 10), bg=C_INPUT, fg=C_TEXT, insertbackground=C_TEXT, relief=tk.FLAT).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6, padx=(0, 8))
    
    def select_f1():
        path = filedialog.askopenfilename(title="Seleccionar Excel Versión 1", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if path:
            f1_var.set(path)
            _log1(f"📁 Archivo 1 cargado: {os.path.basename(path)}\n")

    btn_f1 = tk.Button(wrap1, text="Examinar...", font=(F, 9), bg=C_SURFACE, fg=C_TEXT, relief=tk.FLAT, padx=12, command=select_f1, cursor="hand2")
    btn_f1.pack(side=tk.RIGHT)

    # --- ARCHIVO 2 (Versión 2) ---
    tk.Label(pad1, text="EXCEL VERSIÓN 2 (Nueva / A Comparar)", font=(F, 8, "bold"), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(0, 4))
    f2_var = tk.StringVar()
    wrap2 = tk.Frame(pad1, bg=C_BG)
    wrap2.pack(fill=tk.X, pady=(0, 15))
    
    tk.Entry(wrap2, textvariable=f2_var, font=(F, 10), bg=C_INPUT, fg=C_TEXT, insertbackground=C_TEXT, relief=tk.FLAT).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6, padx=(0, 8))
    
    def select_f2():
        path = filedialog.askopenfilename(title="Seleccionar Excel Versión 2", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if path:
            f2_var.set(path)
            _log1(f"📁 Archivo 2 cargado: {os.path.basename(path)}\n")

    btn_f2 = tk.Button(wrap2, text="Examinar...", font=(F, 9), bg=C_SURFACE, fg=C_TEXT, relief=tk.FLAT, padx=12, command=select_f2, cursor="hand2")
    btn_f2.pack(side=tk.RIGHT)

    def ejecutar_comparacion_excel():
        p1 = f1_var.get()
        p2 = f2_var.get()

        if not p1 or not p2:
            messagebox.showwarning("Archivos faltantes", "Por favor seleccione ambos archivos Excel para comparar.")
            return

        _log1("\n" + "="*50 + "\n")
        _log1("🔍 [INICIANDO AUDITORÍA DETALLADA DE CAMPOS]...\n", "success")

        try:
            df1 = pd.read_excel(p1)
            df2 = pd.read_excel(p2)

            cols1 = list(df1.columns)
            cols2 = list(df2.columns)
            
            _log1(f"• Total de filas Archivo 1: {len(df1)}\n")
            _log1(f"• Total de filas Archivo 2: {len(df2)}\n")

            if cols1 != cols2:
                _log1("⚠️ ADVERTENCIA: Las columnas de ambos archivos no son idénticas exactamente.\n", "warning")

            # Lectura directa sin normalización de texto especial
            df1_norm = df1.astype(str).fillna("")
            df2_norm = df2.astype(str).fillna("")

            # Detectar columna clave (RUC, DNI o NumDocumento)
            possible_keys = ['NUMDOCUMENTO', 'RUC', 'DNI', 'DOCUMENTO', 'NUMDNI']
            key_col = None
            for col in df1.columns:
                if str(col).strip().upper() in possible_keys:
                    key_col = col
                    break
            
            if not key_col:
                key_col = df1.columns[0]

            _log1(f"📌 Columna clave detectada para emparejar registros: '{key_col}'\n", "muted")

            dict1 = {str(row[key_col]): row for _, row in df1_norm.iterrows()}
            dict2 = {str(row[key_col]): row for _, row in df2_norm.iterrows()}

            keys1 = set(dict1.keys())
            keys2 = set(dict2.keys())

            comunes = keys1.intersection(keys2)
            solo_1 = keys1 - keys2
            solo_2 = keys2 - keys1

            diferencias_encontradas = 0

            if solo_1:
                _log1(f"❌ Registros que están en el Archivo 1 pero FALTAN en el Archivo 2 ({len(solo_1)}):\n", "error")
                for k in sorted(solo_1)[:20]:
                    _log1(f"   - Clave / RUC ausente: {k}\n", "error")

            if solo_2:
                _log1(f"❌ Registros EXTRAS que están en el Archivo 2 y no en el Archivo 1 ({len(solo_2)}):\n", "error")
                for k in sorted(solo_2)[:20]:
                    _log1(f"   - Clave / RUC adicional: {k}\n", "error")

            _log1(f"\n🔎 Analizando datos campo por campo en los {len(comunes)} registros coincidentes...\n", "success")
            
            common_cols = [c for c in df1.columns if c in df2.columns]

            for k in sorted(comunes):
                r1 = dict1[k]
                r2 = dict2[k]
                
                cambios_en_registro = []
                for col in common_cols:
                    val1 = str(r1[col])
                    val2 = str(r2[col])
                    if val1 != val2:
                        cambios_en_registro.append((col, val1, val2))

                if cambios_en_registro:
                    diferencias_encontradas += 1
                    _log1(f"\n⚠️ RUC / Documento [{k}] tiene datos incorrectos o modificados:\n", "error")
                    for col, v1, v2 in cambios_en_registro:
                        _log1(f"   • Campo '{col}':\n", "warning")
                        _log1(f"     - Original (Archivo 1) : '{v1}'\n", "muted")
                        _log1(f"     - Modificado (Archivo 2): '{v2}'\n", "error")

            if diferencias_encontradas == 0 and not solo_1 and not solo_2:
                _log1("\n🎉 ¡RESULTADO PERFECTO: Los dos archivos son 100% IDÉNTICOS en todos sus datos y campos!\n", "success")
                messagebox.showinfo("Éxito", "¡Los dos archivos Excel son 100% idénticos en contenido!")
            else:
                _log1(f"\n❌ Auditoría finalizada: Se encontraron {diferencias_encontradas} registros con datos distintos, además de {len(solo_1)} faltantes y {len(solo_2)} sobrantes.\n", "error")
                messagebox.showwarning("Diferencias Encontradas", "Se detectaron datos incorrectos o modificados. Revisa el reporte detallado.")

        except Exception as exc:
            _log1(f"\n❌ Error crítico al procesar los archivos: {exc}\n", "error")
            messagebox.showerror("Error", str(exc))

    btn_run1 = tk.Button(pad1, text="Iniciar Auditoría Detallada de Datos", font=(F, 10, "bold"), bg=C_SUCCESS, fg="#0F172A", relief=tk.FLAT, padx=20, pady=8, command=ejecutar_comparacion_excel, cursor="hand2")
    btn_run1.pack(anchor=tk.W, pady=(0, 15))

    tk.Label(pad1, text="REPORTE DETALLADO DE INCORRECTOS", font=(F, 8, "bold"), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(0, 4))
    log_area1 = scrolledtext.ScrolledText(pad1, font=("Consolas", 9), bg=C_INPUT, fg="#94A3B8", relief=tk.FLAT, padx=12, pady=8, height=14)
    log_area1.pack(fill=tk.BOTH, expand=True)
    log_area1.tag_config("success", foreground=C_SUCCESS)
    log_area1.tag_config("error", foreground=C_ERROR)
    log_area1.tag_config("warning", foreground=C_WARNING)
    log_area1.tag_config("muted", foreground=C_MUTED)

    def _log1(msg, tag=""):
        log_area1.config(state=tk.NORMAL)
        log_area1.insert(tk.END, msg, tag)
        log_area1.see(tk.END)
        log_area1.config(state=tk.DISABLED)


    # ==================== PESTAÑA 2: COMPARAR POR LISTAS (COPIAR Y PEGAR) ====================
    tab2 = tk.Frame(notebook, bg=C_BG)
    notebook.add(tab2, text="  📋 Copiar y Pegar Listas (RUCs / Campos)  ")

    pad2 = tk.Frame(tab2, bg=C_BG)
    pad2.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

    tk.Label(pad2, text="Comparador Directo por Listas de Texto", font=(F, 18, "bold"), bg=C_BG, fg=C_TEXT).pack(anchor=tk.W)
    tk.Label(pad2, text="Pega las columnas (RUCs, Direcciones, Nombres, etc.) de la Lista 1 y Lista 2 para verificar coincidencias.", font=(F, 9), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(2, 10))
    tk.Frame(pad2, height=1, bg=C_BORDER).pack(fill=tk.X, pady=(0, 10))

    text_frames = tk.Frame(pad2, bg=C_BG)
    text_frames.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    text_frames.columnconfigure(0, weight=1)
    text_frames.columnconfigure(1, weight=1)
    text_frames.rowconfigure(0, weight=1)

    frame_left = tk.Frame(text_frames, bg=C_BG)
    frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    tk.Label(frame_left, text="LISTA 1 (Pegar columna base)", font=(F, 8, "bold"), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(0, 4))
    text_list1 = scrolledtext.ScrolledText(frame_left, font=("Consolas", 9), bg=C_INPUT, fg=C_TEXT, insertbackground=C_TEXT, relief=tk.FLAT, height=8)
    text_list1.pack(fill=tk.BOTH, expand=True)

    frame_right = tk.Frame(text_frames, bg=C_BG)
    frame_right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
    tk.Label(frame_right, text="LISTA 2 (Pegar columna a comparar)", font=(F, 8, "bold"), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(0, 4))
    text_list2 = scrolledtext.ScrolledText(frame_right, font=("Consolas", 9), bg=C_INPUT, fg=C_TEXT, insertbackground=C_TEXT, relief=tk.FLAT, height=8)
    text_list2.pack(fill=tk.BOTH, expand=True)

    def ejecutar_comparacion_listas():
        t1_raw = text_list1.get("1.0", tk.END).strip()
        t2_raw = text_list2.get("1.0", tk.END).strip()

        if not t1_raw or not t2_raw:
            messagebox.showwarning("Listas vacías", "Por favor pega datos en ambos cuadros de texto antes de comparar.")
            return

        items1 = [line.strip() for line in t1_raw.splitlines() if line.strip()]
        items2 = [line.strip() for line in t2_raw.splitlines() if line.strip()]

        _log2("\n" + "="*50 + "\n")
        _log2(f"📋 [COMPARANDO LISTAS] | Total en Lista 1: {len(items1)} | Total en Lista 2: {len(items2)}\n", "success")

        c1 = Counter(items1)
        c2 = Counter(items2)

        if c1 == c2:
            _log2("🎉 ¡ÉXITO! Ambas listas son 100% IDÉNTICAS (tienen los mismos elementos con sus respectivas cantidades).\n", "success")
            messagebox.showinfo("Listas Idénticas", "¡Las listas de datos son exactamente iguales!")
        else:
            _log2("⚠️ SE ENCONTRARON DIFERENCIAS ENTRE LAS DOS LISTAS:\n", "warning")
            
            faltantes = c1 - c2  
            sobrantes = c2 - c1  

            if faltantes:
                _log2(f"❌ Elementos que FALTAN en la Lista 2 (o tienen menos repeticiones):\n", "error")
                for item, cant in faltantes.items():
                    _log2(f"   • {item} (Faltan {cant})\n", "error")

            if sobrantes:
                _log2(f"❌ Elementos EXTRAS en la Lista 2 (o tienen más repeticiones):\n", "error")
                for item, cant in sobrantes.items():
                    _log2(f"   • {item} (Sobra(n) {cant})\n", "error")

            messagebox.showwarning("Diferencias", "Se detectaron diferencias en las listas. Revisa el reporte en rojo.")

    def limpiar_listas():
        text_list1.delete("1.0", tk.END)
        text_list2.delete("1.0", tk.END)
        log_area2.config(state=tk.NORMAL)
        log_area2.delete("1.0", tk.END)
        log_area2.config(state=tk.DISABLED)

    btn_box = tk.Frame(pad2, bg=C_BG)
    btn_box.pack(anchor=tk.W, pady=(5, 10))

    btn_run2 = tk.Button(btn_box, text="Comparar Listas Pegadas", font=(F, 10, "bold"), bg=C_SUCCESS, fg="#0F172A", relief=tk.FLAT, padx=20, pady=8, command=ejecutar_comparacion_listas, cursor="hand2")
    btn_run2.pack(side=tk.LEFT, padx=(0, 10))

    btn_clear = tk.Button(btn_box, text="🧹 Limpiar Todo y Borrar", font=(F, 10, "bold"), bg=C_SURFACE, fg=C_TEXT, relief=tk.FLAT, padx=16, pady=8, command=limpiar_listas, cursor="hand2")
    btn_clear.pack(side=tk.LEFT)

    tk.Label(pad2, text="RESULTADOS DE LA COMPARACIÓN DE LISTAS", font=(F, 8, "bold"), bg=C_BG, fg=C_MUTED).pack(anchor=tk.W, pady=(0, 4))
    log_area2 = scrolledtext.ScrolledText(pad2, font=("Consolas", 9), bg=C_INPUT, fg="#94A3B8", relief=tk.FLAT, padx=12, pady=8, height=8)
    log_area2.pack(fill=tk.BOTH, expand=True)
    log_area2.tag_config("success", foreground=C_SUCCESS)
    log_area2.tag_config("error", foreground=C_ERROR)
    log_area2.tag_config("warning", foreground=C_WARNING)

    def _log2(msg, tag=""):
        log_area2.config(state=tk.NORMAL)
        log_area2.insert(tk.END, msg, tag)
        log_area2.see(tk.END)
        log_area2.config(state=tk.DISABLED)

    root.mainloop()

if __name__ == "__main__":
    run_comparator_gui()