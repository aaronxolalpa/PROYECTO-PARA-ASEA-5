import easyocr
from pdf2image import convert_from_path
import pandas as pd
import re
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import subprocess
import platform

class PDFInfoExtractorGUI:
    def __init__(self, master):
        self.master = master
        master.title("Extractor de Información de PDFs")
        master.geometry("900x600")

        # Configuración del extractor
        self.poppler_path = r'C:\Users\sergio\Desktop\PROYECTO_OCR\poppler-24.08.0\Library\bin'
        self.reader = easyocr.Reader(['es'])

        # Lista para almacenar resultados
        self.results_list = []

        # Nuevo: variable para rastrear carpeta de PDFs actual
        self.current_pdf_folder = None
        self.current_pdf_files = []
        self.current_pdf_index = 0
        self.current_pdf_path = None  # Para rastrear el PDF actual
        
        # Para almacenar los datos cargados del Excel
        self.excel_data = None

        # Crear Frame para selección de carpeta de PDFs
        pdf_frame = tk.Frame(master)
        pdf_frame.pack(pady=10, padx=10, fill='x')

        self.pdf_folder_var = tk.StringVar()
        tk.Label(pdf_frame, text="Carpeta de PDFs:").pack(side=tk.LEFT)
        tk.Entry(pdf_frame, textvariable=self.pdf_folder_var, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(pdf_frame, text="Seleccionar Carpeta", command=self.select_pdf_folder).pack(side=tk.LEFT)

        # Progreso de procesamiento
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(master, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(pady=10, padx=10, fill='x')

        # Etiqueta de estado
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(master, textvariable=self.status_var)
        self.status_label.pack(pady=5)

        # Crear campos de entrada
        fields = [
            ("Oficio", 'oficio_var'),
            ("Nombre", 'nombre_var'),
            ("Razón Social", 'razon_social_var'),
            ("Dirección", 'direccion_var'),
            ("Teléfono", 'telefono_var'),
            ("Correo", 'correo_var')
        ]

        # Frame para campos de entrada
        entry_frame = tk.Frame(master)
        entry_frame.pack(pady=10, padx=10, fill='both', expand=True)

        self.entry_vars = {}
        for label_text, var_name in fields:
            # Crear frame para cada campo
            field_frame = tk.Frame(entry_frame)
            field_frame.pack(fill='x', pady=5)
            
            # Label
            tk.Label(field_frame, text=label_text + ":", width=15, anchor='w').pack(side=tk.LEFT)
            
            # Entry
            var = tk.StringVar()
            self.entry_vars[var_name] = var
            entry = tk.Entry(field_frame, textvariable=var, width=50)
            entry.pack(side=tk.LEFT, expand=True, fill='x', padx=5)

        # Botones de acción
        button_frame = tk.Frame(master)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="Procesar Carpeta", command=self.process_pdf_folder).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Agregar a Tabla", command=self.add_to_table).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Guardar Excel", command=self.save_to_excel).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Limpiar", command=self.clear_fields).pack(side=tk.LEFT, padx=5)

        # Tabla de resultados
        self.tree_frame = tk.Frame(master)
        self.tree_frame.pack(pady=10, padx=10, fill='both', expand=True)

        self.tree = ttk.Treeview(self.tree_frame, columns=('Oficio', 'Nombre', 'Razón Social', 'Dirección', 'Teléfono', 'Correo'), show='headings')
        
        # Configurar columnas
        for col in self.tree['columns']:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        self.tree.pack(side='left', fill='both', expand=True)

        # Scrollbar para la tabla
        scrollbar = ttk.Scrollbar(self.tree_frame, orient='vertical', command=self.tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscroll=scrollbar.set)

        # Ruta de guardado de Excel
        self.excel_save_path = r'C:\Users\sergio\Desktop\PROYECTO_OCR\resultados_ocr'
        os.makedirs(self.excel_save_path, exist_ok=True)

        # Variables para el estado de modificación
        self.is_modifying_after_initial_save = False
        self.next_pdf_button = None
        
        # Guardar la ruta del último Excel guardado
        self.last_excel_path = None

    def select_pdf_folder(self):
        """Seleccionar carpeta de PDFs"""
        folder = filedialog.askdirectory(title="Seleccionar carpeta de PDFs")
        if folder:
            self.pdf_folder_var.set(folder)

    def _extract_structured_info(self, lines):
        info = {
            "Nombre": "",
            "Razon Social": "",
            "Dirección": "",
            "Teléfono": "",
            "Correo": ""
        }

        # nombre
        if lines:
            nombre = " ".join(lines[:1]).strip()
            nombre = nombre.replace("C.","").replace("c.","").strip()
            info["Nombre"] = nombre

        # razon social
        company_keywords = ["s.a.", "s.a. de c.v.", "sa de cv", "s.a. de c.v.", "s.a. de c.v", ", s.a. de c.v."]
        
        for i, line in enumerate(lines):
            # Caso especial: si la línea contiene "Empresa", la siguiente línea es la razón social
            if "empresa" in line.lower() and i + 1 < len(lines):
                info["Razon Social"] = lines[i+1].strip()
                break
                
            # Caso original: detección por keywords
            line_lower = line.lower()
            if any(keyword in line_lower for keyword in company_keywords):
                info["Razon Social"] = line.strip()
                
                # Verifica líneas adicionales que podrían formar parte de la razón social
                current_index = i + 1
                while current_index < len(lines):
                    next_line = lines[current_index].strip()
                    next_line_lower = next_line.lower()
                    
                    # Condiciones para no incluir la línea
                    should_exclude = any([
                        any(keyword in next_line_lower for keyword in company_keywords),
                        next_line.startswith("Tel:"),
                        "@" in next_line,
                        re.search(r'\d{4,5}', next_line),
                        len(next_line) < 3,  # Líneas muy cortas probablemente no son parte del nombre
                        re.search(r'rfc|registro|fecha|domicilio|dirección', next_line_lower)
                    ])
                    
                    if should_exclude:
                        break
                    
                    info["Razon Social"] += " " + next_line
                    current_index += 1
                
                break
        # direccion
        address_patterns = [
            r'Prolongación\s+\w+',
            r'colonia\s+\w+',
            r'municipio\s+de\s+\w+', 
            r'estado\s+de\s+\w+',
            r'c\.p\.\s+\d{4,5}'
        ]
        for line in lines:
            for pattern in address_patterns:
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    info["Dirección"] += line + " "
        info["Dirección"] = info["Dirección"].strip()

        # numero de teléfono
        phone_pattern = r'(?:tel(?:éfono)?:?\s*)?(\+?\d{1,2}\s?)?(\(?\d{2,3}\)?[-.\s]?)?\d{3,4}[-.\s]?\d{4}'
        for line in lines:
            match = re.search(phone_pattern, line, re.IGNORECASE)
            if match:
                # Remove "Tel:", "Teléfono:", etc. from the start of the match
                phone = re.sub(r'^.*?:', '', match.group(0)).strip()
                info["Teléfono"] = phone
                break

        # correo electronicos
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        all_emails = []
        
        # Primera pasada: buscar en cada línea individual
        for line in lines:
            email_matches = re.findall(email_pattern, line.lower())  # Convertir a minúsculas para normalizar
            if email_matches:
                # Limpiar cada correo encontrado (eliminar espacios, caracteres no deseados)
                clean_emails = [email.strip() for email in email_matches]
                all_emails.extend(clean_emails)
        
        # Segunda pasada: buscar en las combinaciones de líneas consecutivas
        combined_text = " ".join(lines)
        email_matches_combined = re.findall(email_pattern, combined_text.lower())
        
        # Añadir cualquier correo encontrado en el texto combinado que no esté ya en la lista
        for email in email_matches_combined:
            clean_email = email.strip()
            if clean_email not in all_emails:
                all_emails.append(clean_email)
        
        # Eliminar duplicados y ordenar
        all_emails = list(set(all_emails))
        all_emails.sort()
        
        # Si se encontraron correos, unirlos con comas para el campo Correo
        if all_emails:
            info["Correo"] = ";".join(all_emails)

        return info
    
    def open_current_pdf(self):
        """Abrir el PDF actual con el visor predeterminado del sistema"""
        if self.current_pdf_path and os.path.exists(self.current_pdf_path):
            try:
                if platform.system() == 'Windows':
                    os.startfile(self.current_pdf_path)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.call(['open', self.current_pdf_path])
                else: 
                    subprocess.call(['xdg-open', self.current_pdf_path])
                    
                self.status_var.set(f"PDF abierto: {os.path.basename(self.current_pdf_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el PDF: {str(e)}")
        else:
            messagebox.showwarning("Advertencia", "No hay un PDF actual para abrir")
 
    def add_to_table(self):
        """Agregar datos a la tabla manualmente"""
        data = {
            'Oficio': self.entry_vars['oficio_var'].get(),
            'Nombre': self.entry_vars['nombre_var'].get(),
            'Razón Social': self.entry_vars['razon_social_var'].get(),
            'Dirección': self.entry_vars['direccion_var'].get(),
            'Teléfono': self.entry_vars['telefono_var'].get(),
            'Correo': self.entry_vars['correo_var'].get()
        }

        # Validar que no estén vacíos los campos principales
        if not data['Oficio'] or not data['Nombre']:
            messagebox.showerror("Error", "Oficio y Nombre son campos obligatorios")
            return

        # Agregar a la lista de resultados
        self.results_list.append(data)

        # Agregar a la tabla
        self.tree.insert('', 'end', values=(
            data['Oficio'], 
            data['Nombre'], 
            data['Razón Social'], 
            data['Dirección'], 
            data['Teléfono'], 
            data['Correo']
        ))

        # Si estamos en modo de modificación después de guardado inicial
        if self.is_modifying_after_initial_save:
            # Intentar cargar el siguiente PDF sin limpiar los campos
            result = self._load_next_pdf_data()
            
            # Si no hay más PDFs, terminar el proceso de modificación
            if not result:
                self.is_modifying_after_initial_save = False
                if self.next_pdf_button:
                    self.next_pdf_button.destroy()
                    self.next_pdf_button = None
                messagebox.showinfo("Información", "No hay más PDFs para procesar")
        else:
            # Limpiar campos después de agregar solo si NO estamos en modo modificación
            self.clear_fields()

    def process_pdf_folder(self):
        """Procesar todos los PDFs en la carpeta seleccionada"""
        pdf_folder = self.pdf_folder_var.get()
        if not pdf_folder:
            messagebox.showerror("Error", "Seleccione una carpeta de PDFs primero")
            return

        # Obtener lista de PDFs
        pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showerror("Error", "No se encontraron archivos PDF en la carpeta")
            return

        # Limpiar resultados anteriores
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.results_list.clear()

        # Configurar barra de progreso
        total_pdfs = len(pdf_files)
        self.progress_var.set(0)

        # Procesar cada PDF
        for idx, pdf_filename in enumerate(pdf_files, 1):
            try:
                # Ruta completa del PDF
                pdf_path = os.path.join(pdf_folder, pdf_filename)
                
                # Actualizar estado
                self.status_var.set(f"Procesando PDF {idx} de {total_pdfs}: {pdf_filename}")
                self.master.update_idletasks()

                # pdf a imagen
                pages = convert_from_path(pdf_path, dpi=300, poppler_path=self.poppler_path)
                first_page = pages[0]

                # coordenadas de recorte (ajustar según sea necesario)
                crop_coords = (0, 720, 1750, 1550)
                first_page = first_page.crop(crop_coords)
                first_page.save("temp_ocr_image.png")

                # usar OCR
                results = self.reader.readtext("temp_ocr_image.png", detail=0)
                
                # procesar lineas
                lines = [line.strip() for line in results if line.strip()]
                
                # extraer información
                info = self._extract_structured_info(lines)
                
                # agregar nombre de archivo como oficio
                info["Oficio"] = os.path.splitext(pdf_filename)[0].replace("_","/")
               
                # limpiar archivos temporales
                os.remove("temp_ocr_image.png")

                # agregar a resultados
                data = {
                    'Oficio': info.get('Oficio', ''),
                    'Nombre': info.get('Nombre', ''),
                    'Razón Social': info.get('Razon Social', ''),
                    'Dirección': info.get('Dirección', ''),
                    'Teléfono': info.get('Teléfono', ''),
                    'Correo': info.get('Correo', '')
                }

                # Agregar a la lista de resultados
                self.results_list.append(data)

                # Agregar a la tabla
                self.tree.insert('', 'end', values=(
                    data['Oficio'], 
                    data['Nombre'], 
                    data['Razón Social'], 
                    data['Dirección'], 
                    data['Teléfono'], 
                    data['Correo']
                ))

                # Actualizar barra de progreso 
                self.progress_var.set((idx / total_pdfs) * 100)
                self.master.update_idletasks()

            except Exception as e:
                messagebox.showwarning("Advertencia", f"No se pudo procesar {pdf_filename}: {str(e)}")

        # Finalizar procesamiento
        self.status_var.set(f"Procesamiento completado. {total_pdfs} PDFs analizados.")
        messagebox.showinfo("Completado", f"Se procesaron {total_pdfs} PDFs.")

        if not self.results_list:
            messagebox.showerror("Error", "No hay datos para guardar")
            return

        # Generar nombre de archivo con fecha y hora
        excel_filename = f"resultados_OCR_sin_modificacion.xlsx"
        excel_path = os.path.join(self.excel_save_path, excel_filename)

        try:
            # Convertir a DataFrame y guardar
            df = pd.DataFrame(self.results_list)
            
            # Asegurar que los valores NaN se conviertan a cadenas vacías
            df = df.fillna('')
            
            df.to_excel(excel_path, index=False)
            
            # Guardar la ruta del Excel para uso posterior
            self.last_excel_path = excel_path
            
            # Primero mostrar mensaje de éxito de guardado
            messagebox.showinfo("Éxito", f"Datos guardados en {excel_path}")
            
            # Luego preguntar si desea hacer modificaciones
            respuesta = messagebox.askyesno("Modificaciones", 
                "¿Desea hacer modificaciones a los datos?\n"
                "Si selecciona 'Sí', se procesarán los PDFs secuencialmente.")
            
            if respuesta:
                # Guardar la carpeta de PDFs actual si aún no se ha guardado
                if not self.current_pdf_folder:
                    self.current_pdf_folder = self.pdf_folder_var.get()
                    self.current_pdf_files = [f for f in os.listdir(self.current_pdf_folder) 
                                               if f.lower().endswith('.pdf')]
                    self.current_pdf_index = 0  # Inicializar el índice
                
                # Cargar los datos del Excel
                self.excel_data = pd.read_excel(self.last_excel_path)
                # Reemplazar valores NaN con cadenas vacías
                self.excel_data = self.excel_data.fillna('')
                
                # Limpiar la tabla actual
                for i in self.tree.get_children():
                    self.tree.delete(i)
                self.results_list.clear()

                # Marcar que estamos en modo de modificación
                self.is_modifying_after_initial_save = True

                # Intentar cargar el primer PDF
                if self._load_next_pdf_data():
                    # Crear botón para cargar siguiente PDF
                    if self.next_pdf_button:
                        self.next_pdf_button.destroy()
                    self.next_pdf_button = tk.Button(self.master, text="Siguiente PDF", command=self._load_next_pdf_data)
                    self.next_pdf_button.pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")

    def _load_next_pdf_data(self):
        """Cargar datos del siguiente PDF desde Excel"""
        if not self.current_pdf_files or self.current_pdf_index >= len(self.current_pdf_files):
            messagebox.showinfo("Información", "No hay más PDFs para procesar")
            # Reiniciar el índice y limpiar todo
            if self.next_pdf_button:
                self.next_pdf_button.destroy()
                self.next_pdf_button = None
            self.current_pdf_index = 0
            self.current_pdf_folder = None
            self.current_pdf_files = []
            self.current_pdf_path = None
            self.is_modifying_after_initial_save = False
            return False

        # Obtener el PDF actual
        pdf_filename = self.current_pdf_files[self.current_pdf_index]
        pdf_path = os.path.join(self.current_pdf_folder, pdf_filename)
        
        try:
            # Guardar la ruta del PDF actual para poder abrirlo después
            self.current_pdf_path = pdf_path
            
            # Actualizar estado
            self.status_var.set(f"Procesando PDF: {pdf_filename}")
            
            # Obtener el oficio del nombre del archivo
            oficio = os.path.splitext(pdf_filename)[0].replace("_","/")
            
            # Buscar datos en el Excel para este oficio
            if self.excel_data is not None:
                # Buscar el registro con el mismo oficio
                row = self.excel_data[self.excel_data['Oficio'] == oficio]
                
                if not row.empty:
                    # Cargar información desde Excel en los campos de texto
                    self.entry_vars['oficio_var'].set(row['Oficio'].values[0])
                    self.entry_vars['nombre_var'].set(row['Nombre'].values[0])
                    self.entry_vars['razon_social_var'].set(row['Razón Social'].values[0])
                    self.entry_vars['direccion_var'].set(row['Dirección'].values[0])
                    self.entry_vars['telefono_var'].set(row['Teléfono'].values[0])
                    self.entry_vars['correo_var'].set(row['Correo'].values[0])
                else:
                    # Si no encuentra datos en Excel, cargar solo el oficio
                    self.entry_vars['oficio_var'].set(oficio)
                    self.entry_vars['nombre_var'].set('')
                    self.entry_vars['razon_social_var'].set('')
                    self.entry_vars['direccion_var'].set('')
                    self.entry_vars['telefono_var'].set('')
                    self.entry_vars['correo_var'].set('')
            
            # Abrir automáticamente el PDF
            self.open_current_pdf()
            
            # Incrementar el índice para el próximo PDF
            self.current_pdf_index += 1

            return True

        except Exception as e:
            messagebox.showwarning("Advertencia", f"No se pudo cargar los datos del PDF {pdf_filename}: {str(e)}")
            # Incrementar el índice para el próximo PDF aún si hubo error
            self.current_pdf_index += 1
            return False

    def save_to_excel(self):
        """Guardar datos de la tabla en Excel con opción de modificación"""
        if not self.results_list:
            messagebox.showerror("Error", "No hay datos para guardar")
            return

        # Generar nombre de archivo con fecha y hora
        excel_filename = f"resultados_OCR_modificado.xlsx"
        excel_path = os.path.join(self.excel_save_path, excel_filename)

        try:
            # Convertir a DataFrame y guardar
            df = pd.DataFrame(self.results_list)
            
            # Asegurar que los valores NaN se conviertan a cadenas vacías
            df = df.fillna('')
            
            df.to_excel(excel_path, index=False)
            
            # Guardar la ruta del Excel para uso posterior
            self.last_excel_path = excel_path
            
            # Primero mostrar mensaje de éxito de guardado
            messagebox.showinfo("Éxito", f"Datos guardados en {excel_path}")
            
            # Luego preguntar si desea hacer modificaciones
            respuesta = messagebox.askyesno("Modificaciones", 
                "¿Desea hacer modificaciones a los datos?\n"
                "Si selecciona 'Sí', se procesarán los PDFs secuencialmente.")
            
            if respuesta:
                # Guardar la carpeta de PDFs actual si aún no se ha guardado
                if not self.current_pdf_folder:
                    self.current_pdf_folder = self.pdf_folder_var.get()
                    self.current_pdf_files = [f for f in os.listdir(self.current_pdf_folder) 
                                               if f.lower().endswith('.pdf')]
                    self.current_pdf_index = 0  # Inicializar el índice
                
                # Cargar los datos del Excel
                self.excel_data = pd.read_excel(self.last_excel_path)
                # Reemplazar valores NaN con cadenas vacías
                self.excel_data = self.excel_data.fillna('')
                
                # Limpiar la tabla actual
                for i in self.tree.get_children():
                    self.tree.delete(i)
                self.results_list.clear()

                # Marcar que estamos en modo de modificación
                self.is_modifying_after_initial_save = True

                # Intentar cargar el primer PDF
                if self._load_next_pdf_data():
                    # Crear botón para cargar siguiente PDF
                    if self.next_pdf_button:
                        self.next_pdf_button.destroy()
                    self.next_pdf_button = tk.Button(self.master, text="Siguiente PDF", command=self._load_next_pdf_data)
                    self.next_pdf_button.pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")

    def clear_fields(self):
    
        # Limpiar todos los campos de entrada
        for var in self.entry_vars.values():
            var.set('')
    
        # Limpiar etiqueta de estado y barra de progreso
        self.status_var.set('')
        self.progress_var.set(0)
    
        # Limpiar la tabla 
        for item in self.tree.get_children():
            self.tree.delete(item)
    
        # Limpiar la lista de resultados
        self.results_list.clear()
    
        # Si estamos procesando un PDF, reiniciar el estado actual del PDF
        if not self.is_modifying_after_initial_save:
            self.current_pdf_path = None
        
        # Mostrar mensaje de campos limpiados
        self.status_var.set('Campos limpiados')

def main():
    root = tk.Tk()
    app = PDFInfoExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()