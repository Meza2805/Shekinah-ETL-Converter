import threading
import os
import ctypes
import tkinter as tk
from typing import Tuple, Optional, Dict

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np

# --- CONFIGURACIÓN DE IDENTIDAD DE WINDOWS ---
try:
    myappid = 'shekinah.etl.converter.ultimate.2.1'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass
# ------------------------------------------------

class ExcelToSQLApp:
    def __init__(self, root: ttk.Window):
        self.root = root
        self.style = ttk.Style(theme="cosmo")
        self.root.title("Excel to SQL Converter | Shekinah Services")
        
        # Maximizar ventana
        try:
            self.root.state('zoomed')
        except:
            self._center_window(1000, 800)

        self._setup_icons()
        self.root.protocol("WM_DELETE_WINDOW", self.confirm_exit_custom)

        # Variables
        self.file_path = ttk.StringVar()
        self.table_name = ttk.StringVar(value="TempTable")
        
        self.db_options = [
            "🛢️ SQL Server", 
            "🐬 MySQL", 
            "🐘 PostgreSQL", 
            "🪶 SQLite",
            "🔮 Oracle"
        ]
        self.db_type = ttk.StringVar(value=self.db_options[0])
        
        self.status_msg = ttk.StringVar(value="Esperando archivo...")
        self.df: Optional[pd.DataFrame] = None
        
        # Variables de estado del script
        self.full_create_sql: str = ""
        self.full_insert_sql: str = ""
        self.is_generated: bool = False # Bandera para validar
        
        self.logo_img_ref = None
        self.exit_img_ref = None
        self.icon_img_ref = None

        self._init_ui()

    def _center_window(self, width: int, height: int) -> None:
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _setup_icons(self):
        if os.path.exists("icono.ico"):
            try:
                self.icon_img_ref = tk.PhotoImage(file="icono.ico")
                self.root.iconphoto(True, self.icon_img_ref)
                return
            except: pass
        if os.path.exists("icono.ico"):
            try: self.root.iconbitmap("icono.ico")
            except: pass

    def _init_ui(self) -> None:
        main_container = ttk.Frame(self.root, padding=20)
        main_container.pack(fill=BOTH, expand=YES)

        # Header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=X, pady=(0, 20))

        if os.path.exists("shekinah_logo.png"):
            try:
                self.logo_img_ref = tk.PhotoImage(file="shekinah_logo.png")
                while self.logo_img_ref.width() > 250 or self.logo_img_ref.height() > 100:
                    self.logo_img_ref = self.logo_img_ref.subsample(2, 2) 
                lbl_logo = ttk.Label(header_frame, image=self.logo_img_ref)
                lbl_logo.pack(side=LEFT, padx=(0, 15))
            except: pass 

        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=LEFT, fill=X)
        ttk.Label(title_frame, text="Shekinah Services ETL Tool", font=("Segoe UI", 20, "bold"), bootstyle="primary").pack(anchor=W)
        ttk.Label(title_frame, text="Herramienta de Migración de Datos Excel a SQL", font=("Segoe UI", 11), bootstyle="secondary").pack(anchor=W)

        # Paneles
        lbl_frame_top = ttk.Labelframe(main_container, text=" 1. Origen de Datos ", padding=15, bootstyle="primary")
        lbl_frame_top.pack(fill=X, pady=(0, 15))
        
        col_frame = ttk.Frame(lbl_frame_top)
        col_frame.pack(fill=X)
        btn_browse = ttk.Button(col_frame, text="📂 Seleccionar Excel", command=self.browse_file, bootstyle="secondary")
        btn_browse.pack(side=LEFT, padx=(0, 10))
        
        # Entry readonly para mostrar la ruta
        self.entry_file = ttk.Entry(col_frame, textvariable=self.file_path, state="readonly")
        self.entry_file.pack(side=LEFT, fill=X, expand=YES)
        
        # Botón Limpiar Manual
        btn_reset = ttk.Button(col_frame, text="🔄 Nuevo Proceso", command=self.reset_app, bootstyle="light")
        btn_reset.pack(side=RIGHT, padx=(10, 0))

        lbl_frame_conf = ttk.Labelframe(main_container, text=" 2. Configuración SQL ", padding=15, bootstyle="info")
        lbl_frame_conf.pack(fill=X, pady=(0, 15))
        
        grid_frame = ttk.Frame(lbl_frame_conf)
        grid_frame.pack(fill=X)
        ttk.Label(grid_frame, text="Motor Base Datos:").grid(row=0, column=0, sticky=W, padx=5)
        self.combo_db = ttk.Combobox(grid_frame, textvariable=self.db_type, state="readonly", width=22)
        self.combo_db['values'] = self.db_options
        self.combo_db.current(0)
        self.combo_db.grid(row=0, column=1, sticky=W, padx=5)
        
        ttk.Label(grid_frame, text="Nombre Tabla Destino:").grid(row=0, column=2, sticky=W, padx=(20, 5))
        self.entry_table = ttk.Entry(grid_frame, textvariable=self.table_name, width=25)
        self.entry_table.grid(row=0, column=3, sticky=W, padx=5)
        
        btn_gen = ttk.Button(lbl_frame_conf, text="⚡ GENERAR SCRIPTS SQL", command=self.start_generation, bootstyle="success")
        btn_gen.pack(fill=X, pady=15)

        lbl_frame_res = ttk.Labelframe(main_container, text=" 3. Resultados Generados ", padding=10, bootstyle="secondary")
        lbl_frame_res.pack(fill=BOTH, expand=YES)
        
        self.notebook = ttk.Notebook(lbl_frame_res, bootstyle="light")
        self.notebook.pack(fill=BOTH, expand=YES, pady=5)
        self.txt_create = scrolledtext.ScrolledText(self.notebook, height=10, font=("Consolas", 10))
        self.notebook.add(self.txt_create, text="DDL (Create Table)")
        self.txt_insert = scrolledtext.ScrolledText(self.notebook, height=10, font=("Consolas", 10))
        self.notebook.add(self.txt_insert, text="DML (Insert Data)")

        footer_frame = ttk.Frame(main_container, padding=(0, 10, 0, 0))
        footer_frame.pack(fill=X)
        lbl_status = ttk.Label(footer_frame, textvariable=self.status_msg, font=("Segoe UI", 9, "italic"), bootstyle="secondary")
        lbl_status.pack(side=LEFT)
        btn_save = ttk.Button(footer_frame, text="💾 Guardar .SQL", command=self.save_file, bootstyle="primary-outline")
        btn_save.pack(side=RIGHT, padx=5)
        btn_copy = ttk.Button(footer_frame, text="📋 Copiar", command=self.copy_to_clipboard, bootstyle="secondary-outline")
        btn_copy.pack(side=RIGHT, padx=5)

    # --- LÓGICA DE VALIDACIÓN Y LIMPIEZA ---

    def reset_app(self):
        """Limpia toda la interfaz para un nuevo proceso."""
        self.file_path.set("")
        self.table_name.set("TempTable")
        self.status_msg.set("Esperando archivo...")
        
        self.txt_create.delete(1.0, END)
        self.txt_insert.delete(1.0, END)
        
        self.df = None
        self.full_create_sql = ""
        self.full_insert_sql = ""
        self.is_generated = False
        
        # Resetear color del entry de archivo si hubo error
        self.entry_file.configure(bootstyle="default")

    def confirm_exit_custom(self):
        exit_window = ttk.Toplevel(self.root)
        exit_window.title("Confirmar Salida")
        if self.icon_img_ref:
            exit_window.iconphoto(True, self.icon_img_ref)
        elif os.path.exists("icono.ico"):
            try: exit_window.iconbitmap("icono.ico")
            except: pass
        
        w, h = 420, 180
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        exit_window.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        exit_window.resizable(False, False)
        exit_window.transient(self.root)
        exit_window.grab_set()

        content_frame = ttk.Frame(exit_window, padding=20)
        content_frame.pack(fill=BOTH, expand=YES)

        if os.path.exists("exit_image.png"):
            try:
                self.exit_img_ref = tk.PhotoImage(file="exit_image.png")
                while self.exit_img_ref.width() > 100:
                    self.exit_img_ref = self.exit_img_ref.subsample(2, 2)
                lbl_img = ttk.Label(content_frame, image=self.exit_img_ref)
                lbl_img.pack(side=LEFT, padx=(0, 20))
            except: pass

        msg_frame = ttk.Frame(content_frame)
        msg_frame.pack(side=LEFT, fill=BOTH, expand=YES)
        ttk.Label(msg_frame, text="¿Desea cerrar la aplicación?", font=("Segoe UI", 12, "bold")).pack(pady=(5, 5), anchor=W)
        ttk.Label(msg_frame, text="Cualquier script no guardado se perderá.", font=("Segoe UI", 9)).pack(pady=(0, 20), anchor=W)
        btn_frame = ttk.Frame(msg_frame)
        btn_frame.pack(fill=X)
        ttk.Button(btn_frame, text="Salir", bootstyle="danger", command=self.root.destroy).pack(side=RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancelar", bootstyle="secondary", command=exit_window.destroy).pack(side=RIGHT, padx=5)

    def browse_file(self) -> None:
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename:
            self.file_path.set(filename)
            self.status_msg.set("⏳ Analizando archivo...")
            self.entry_file.configure(bootstyle="default") # Reset style
            threading.Thread(target=self._load_excel_thread, args=(filename,), daemon=True).start()

    def _load_excel_thread(self, filename: str) -> None:
        try:
            self.df = pd.read_excel(filename)
            rows, cols = self.df.shape
            self.status_msg.set(f"✅ Archivo cargado: {rows} filas detectadas.")
            clean_name = os.path.splitext(os.path.basename(filename))[0].replace(" ", "_")
            self.table_name.set(clean_name)
            self.is_generated = False # Reset flag si cargan nuevo archivo
        except Exception as e:
            self.status_msg.set("❌ Error de lectura.")
            err_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror("Error I/O", f"No se pudo leer el archivo:\n{err_msg}", parent=self.root, icon='error'))

    def start_generation(self) -> None:
        if self.df is None:
            messagebox.showwarning("Validación", "Por favor seleccione un archivo Excel primero.", parent=self.root, icon='warning')
            self.entry_file.configure(bootstyle="danger") # Resaltar en rojo
            return

        self.status_msg.set("⚙️ Analizando tipos de datos y generando...")
        self.txt_create.delete(1.0, END)
        self.txt_insert.delete(1.0, END)

        try:
            create_sql, insert_sql = self._build_sql()
            self.full_create_sql = create_sql
            self.full_insert_sql = insert_sql
            self.is_generated = True

            self.txt_create.insert(END, create_sql)
            preview = insert_sql[:5000] + "\n\n/* ... (Script truncado para visualización) ... */" if len(insert_sql) > 5000 else insert_sql
            self.txt_insert.insert(END, preview)
            
            messagebox.showinfo("Éxito", "✅ Scripts generados exitosamente.\nRevise la vista previa y guarde el archivo.", parent=self.root, icon='info')
            self.status_msg.set("✅ Generación completada. Listo para guardar.")
            
        except Exception as e:
            self.status_msg.set("❌ Error.")
            messagebox.showerror("Error SQL", f"Fallo al construir query:\n{str(e)}", parent=self.root, icon='error')

    def _detect_column_type(self, series: pd.Series) -> str:
        clean_series = series.dropna()
        if clean_series.empty: return "string"
        numeric_conversion = pd.to_numeric(clean_series, errors='coerce')
        if numeric_conversion.isna().any(): return "string"
        if (numeric_conversion % 1 == 0).all(): return "int"
        return "float"

    def _build_sql(self) -> Tuple[str, str]:
        df = self.df.copy()
        table = self.table_name.get()
        db_selection = self.db_type.get()
        
        full_table_name = table
        create_prefix = ""
        table_suffix = ""
        
        if "SQL Server" in db_selection:
            full_table_name = f"#{table}"
            create_prefix = f"CREATE TABLE {full_table_name}"
        elif "MySQL" in db_selection:
            create_prefix = f"CREATE TEMPORARY TABLE {full_table_name}"
        elif "PostgreSQL" in db_selection:
            create_prefix = f"CREATE TEMP TABLE {full_table_name}"
        elif "SQLite" in db_selection:
            create_prefix = f"CREATE TEMP TABLE {full_table_name}"
        elif "Oracle" in db_selection:
            create_prefix = f"CREATE GLOBAL TEMPORARY TABLE {full_table_name}"
            table_suffix = " ON COMMIT PRESERVE ROWS"

        cols_def = []
        col_types_map = {} 

        for col in df.columns:
            safe_col = str(col).strip().replace(" ", "_").replace(".", "")
            detected_type = self._detect_column_type(df[col])
            col_types_map[col] = detected_type
            
            sql_type = "VARCHAR(255)" 
            if detected_type == "int":
                if "SQLite" in db_selection: sql_type = "INTEGER"
                elif "Oracle" in db_selection: sql_type = "NUMBER(10)"
                else: sql_type = "INT"
            elif detected_type == "float":
                if "SQLite" in db_selection: sql_type = "REAL"
                elif "Oracle" in db_selection: sql_type = "NUMBER(18,4)"
                else: sql_type = "DECIMAL(18,4)"
            else:
                if "SQL Server" in db_selection: sql_type = "NVARCHAR(MAX)"
                elif "Oracle" in db_selection: sql_type = "VARCHAR2(255)"
                else: sql_type = "TEXT"
            
            col_fmt = f"    [{safe_col}] {sql_type}" if "SQL Server" in db_selection else f"    {safe_col} {sql_type}"
            cols_def.append(col_fmt)

        create_script = f"{create_prefix} (\n" + ",\n".join(cols_def) + f"\n){table_suffix};\n"

        inserts = []
        def is_sql_null(val):
            if val is None: return True
            if pd.isna(val): return True
            s_val = str(val).strip().lower()
            if s_val == 'nan' or s_val == 'nat' or s_val == '': return True
            return False

        # Usar 'map' para iterar es mas rapido que where en dataframes grandes
        # Convertimos todo el DF a objetos Python nativos para iteración segura
        df_clean = df.astype(object).where(pd.notnull(df), None)

        for _, row in df_clean.iterrows():
            vals = []
            for col in df.columns:
                val = row[col]
                dtype = col_types_map[col]

                if is_sql_null(val):
                    vals.append("NULL")
                else:
                    if dtype == "string":
                        clean_val = str(val).replace("'", "''")
                        if "Oracle" in db_selection and isinstance(val, pd.Timestamp):
                             vals.append(f"TO_DATE('{clean_val}', 'YYYY-MM-DD HH24:MI:SS')")
                        else:
                            vals.append(f"'{clean_val}'")
                    else:
                        vals.append(str(val))
            
            row_str = "(" + ", ".join(vals) + ")"
            inserts.append(f"INSERT INTO {full_table_name} VALUES {row_str};")

        return create_script, "\n".join(inserts)

    def save_file(self) -> None:
        # VALIDACIÓN: Si no hay script generado, mostrar alerta y salir
        if not self.is_generated or not self.full_create_sql:
            messagebox.showwarning("Atención", "Primero debes generar el script antes de guardar.", parent=self.root, icon='warning')
            return

        f = filedialog.asksaveasfilename(defaultextension=".sql", filetypes=[("SQL Script", "*.sql")], initialfile=f"script_{self.table_name.get()}.sql")
        if f:
            try:
                with open(f, "w", encoding="utf-8") as file:
                    file.write(f"-- Generado por Shekinah ETL | {self.db_type.get()}\n")
                    file.write(self.full_create_sql + "\n")
                    file.write(self.full_insert_sql)
                
                # Mensaje de éxito
                messagebox.showinfo("Guardado", "Archivo exportado correctamente.\n\nLa aplicación se limpiará para un nuevo proceso.", parent=self.root, icon='info')
                
                # LIMPIEZA AUTOMÁTICA
                self.reset_app()
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{str(e)}", parent=self.root, icon='error')

    def copy_to_clipboard(self) -> None:
        if not self.full_create_sql: 
            messagebox.showwarning("Atención", "No hay contenido para copiar.", parent=self.root, icon='warning')
            return
        current_tab = self.notebook.index(self.notebook.select())
        content = self.full_create_sql if current_tab == 0 else self.full_insert_sql
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status_msg.set("Contenido copiado.")

if __name__ == "__main__":
    app_window = ttk.Window(themename="cosmo") 
    app = ExcelToSQLApp(app_window)
    app_window.mainloop()