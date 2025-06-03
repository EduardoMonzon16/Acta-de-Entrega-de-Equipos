from flask import Flask, render_template, request, send_file, redirect, url_for, session, flash
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from io import BytesIO
from mysql.connector import Error
import os
import mysql.connector
import hashlib 


app = Flask(__name__)
app.secret_key = 'clave_secreta_segura'

def conectar():
    return mysql.connector.connect(
        host='localhost',          # o la IP si está en otra máquina
        user='root',   # usuario con permisos para acceder a la BD
        password='Universitario12#',
        database='systembd'     # nombre de la base de datos
    )

# Función para validar credenciales
def validar_credenciales(usuario, password):
    try:
        conexion = conectar()
        cursor = conexion.cursor()
        sql = "SELECT Contraseña FROM usuarios WHERE Nombre = %s"
        cursor.execute(sql, (usuario,))
        resultado = cursor.fetchone()
        cursor.close()
        conexion.close()

        if resultado:
            contraseña_en_bd = resultado[0]
            return contraseña_en_bd == password 
        else:
            return False
    except Error as e:
        print(f"Error en la conexión: {e}")
        return False

# Aplica color de fondo a una celda
def sombrear_celda(celda, color_hex="D9D9D9"):
    tc = celda._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)
    
# Establece la fuente y el tamaño del texto en una celda.
def aplicar_fuente_celda(celda, fuente="Calibri", tam=11):
    for parrafo in celda.paragraphs:
        for run in parrafo.runs:
            run.font.name = fuente
            run.font.size = Pt(tam)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), fuente)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form['usuario']
        password = request.form['password']

        if validar_credenciales(usuario, password):
            session['usuario'] = usuario
            return redirect(url_for('formulario'))
        else:
            flash('Usuario o contraseña incorrectos', 'error')

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/', methods=['GET', 'POST'])
def formulario():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    fecha_actual = datetime.today().strftime('%Y-%m-%d')

    if request.method == 'POST':
        # Datos personales
        nombre = request.form['nombre']
        correo_usuario = request.form['correo'].strip()  
        correo = f"{correo_usuario}@forvismazars.com"
        cargo = request.form['cargo']
        usuario = request.form['usuario']
        telefono = request.form['telefono']

        # Hardware
        tipo = request.form['tipo']
        marca = request.form['marca']
        modelo = request.form['modelo']
        serial = request.form['serial']
        procesador = request.form['procesador']
        ram = request.form['ram']
        disco = request.form['disco']
        perifericos = request.form['perifericos']

        # Entrega de equipo
        nombre_recibe = request.form.get('nombre_recibe', '')
        fecha_recibe = request.form.get('fecha_recibe', '')
        nombre_entrega = request.form.get('nombre_entrega', '')
        fecha_entrega = request.form.get('fecha_entrega', '')

        # Datos del equipo
        fecha_compra = request.form.get('fecha_compra', '')
        equipo = request.form.get('equipo', '')
        marca_equipo = request.form.get('marca_equipo', '')
        hostname = request.form.get('hostname', '')
        modelo_equipo = request.form.get('modelo_equipo', '')
        detalle = request.form.get('detalle', '')
        serie_equipo = request.form.get('serie_equipo', '')
        os_equipo = request.form.get('os', '')
        garantia_raw = request.form.get('garantia', '')

        # Historial de usuarios (listas)
        historial_inicio = request.form.getlist('historial_inicio[]')
        historial_fin = request.form.getlist('historial_fin[]')
        historial_usuario = request.form.getlist('historial_usuario[]')

        # Historial de eventos
        evento_fechas = request.form.getlist('evento_fecha[]')
        evento_observaciones = request.form.getlist('evento_observaciones[]')

        # Agregar formato 'dd-mm-aaaa' en fecha-compra
        try:
            fecha_compra_formateada = datetime.strptime(fecha_compra, "%Y-%m-%d").strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            fecha_compra_formateada = fecha_compra

        # Agregar formato 'dd-mm-aaaa' en garantia
        try:
            garantia_formateada = datetime.strptime(garantia_raw, "%Y-%m-%d").strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            garantia_formateada = garantia_raw or ""

        # Agregar formato 'dd-mm-aaaa' en fecha-recibe
        try:
            fecha_recibe_formateada = datetime.strptime(fecha_recibe, "%Y-%m-%d").strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            fecha_recibe_formateada = fecha_recibe

        # Agregar formato 'dd-mm-aaaa' en fecha-entrega
        try:
            fecha_entrega_formateada = datetime.strptime(fecha_entrega, "%Y-%m-%d").strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            fecha_entrega_formateada = fecha_entrega

        # Crea un nombre único para el documento con el nombre y la fecha/hora
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"Acta_Entrega_{nombre.replace(' ', '_')}_{timestamp}.docx"

        # Mantenimiento de hardware
        mantenimiento_campos = [
            'danos_fisicos',
            'teclado',
            'raton',
            'cargador',
            'puertos',
            'camara',
            'microfono',
            'limpieza_interna',
            'limpieza_externa',
            'parlantes',
            'pantalla',
            'bateria',
            'ethernet',
            'bluetooth',
            'wifi'
        ]

        # Datos del usuario
        datos = {
            "Nombre:": nombre,
            "Correo:": correo,
            "Cargo:": cargo,
            "Usuario de red:": usuario,
            "Teléfono:": telefono
        }

        # Observaciones
        texto_obs = (
            "Certifico que los elementos detallados en el presente documento me han sido entregados para mi cuidado "
            "y custodia con el propósito de cumplir con las tareas y asignaciones propias de mi cargo, siendo estos "
            "de mi única y exclusiva responsabilidad. Me comprometo a usar correctamente los recursos solo para los "
            "fines establecidos, y a no instalar ni permitir la instalación de software por personal ajeno al personal "
            "de TI de Forvis Mazars Perú. De igual forma me comprometo a devolver el equipo en las mismas condiciones "
            "y con los mismos accesorios que me fue entregado, cuando se me programe algún cambio de equipo o el "
            "vínculo laboral haya culminado."
        )

        # Datos del equipo
        datos_equipo = {
            "Fecha de compra:": fecha_compra_formateada,
            "Equipo:": equipo,
            "Marca:": marca_equipo,
            "Hostname:": hostname,
            "Modelo:": modelo_equipo,
            "Detalle:": detalle,
            "Serie:": serie_equipo,
            "Sistema Operativo:": os_equipo,
            "Garantía:": garantia_formateada,
        }

        # Preguntas para mantenimiento de hardware
        preguntas = [
            ('¿El equipo presenta daños físicos?', 'danos_fisicos'),
            ('¿El teclado funciona correctamente?', 'teclado'),
            ('¿El ratón/trackpad funciona correctamente?', 'raton'),
            ('¿El cargador se encuentra en buen estado?', 'cargador'),
            ('¿Los puertos funcionan correctamente?', 'puertos'),
            ('¿La cámara funciona correctamente?', 'camara'),
            ('¿El micrófono funciona correctamente?', 'microfono'),
            ('¿El equipo recibió limpieza interna?', 'limpieza_interna'),
            ('¿El equipo recibió limpieza externa?', 'limpieza_externa'),
            ('¿Los parlantes funcionan correctamente?', 'parlantes'),
            ('¿La pantalla funciona correctamente?', 'pantalla'),
            ('¿La batería funciona correctamente?', 'bateria'),
            ('¿La tarjeta ethernet funciona correctamente?', 'ethernet'),
            ('¿El Bluetooth funciona correctamente?', 'bluetooth'),
            ('¿La tarjeta wifi funciona correctamente?', 'wifi'),
        ]

        # Preguntas de mantenimiento de software
        preguntas_software = [
            ("¿Se verificaron los programas vigentes?", "programas_vigentes"),
            ("¿Se eliminaron los programas externos?", "programas_externos"),
            ("¿Se realizó la limpieza de archivos temporales?", "limpieza_temporal"),
            ("¿Se eliminaron los perfiles antiguos?", "perfiles_antiguos"),
            ("¿Se realizaron actualizaciones de Windows?", "actualizaciones_windows"),
            ("¿Se comprobó el estado del disco duro?", "estado_disco"),
            ("¿Se realizo el backup del usuario anterior?", "backup_usuario"),
            ("¿Se realizó desfragmentación del disco duro?", "desfragmentacion"),
        ]

        # Lista de programas del usuario
        programas_por_columna = [
            ["Anydesk", "Office 365", "Cisco VPN", "PDF24", "Microsoft Teams", "Microsoft Defender", "Acrobat Reader"],
            ["Atlas", "Auditsoft", "---", "---", "---", "---", "---"],
            ["Concar", "Starsoft", "PDT", "PLAME", "---", "---", "---"],
            ["Impresoras", "Scanners", "---", "---", "---", "---", "---"],
            ["PDT", "PLAME", "PLE", "Renta Anual", "Mis declaraciones", "PDB", "---"]
        ]

        mantenimiento_data = {}
        for campo in mantenimiento_campos:
            estado_str = request.form.get(f"{campo}_sn", "false")
            estado_bool = estado_str.lower() == "true"
            detalle = request.form.get(f"{campo}_detalle", "")
            
            mantenimiento_data[campo] = {
                "estado": estado_bool,
                "detalle": detalle
            }

        doc = Document()

        # Encabezado con logo Forvis Mazars
        section = doc.sections[0]
        header = section.header
        for p in header.paragraphs:
            p.clear()
        header_paragraph = header.add_paragraph()
        header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        logo_path = "static/logo.png"
        if os.path.exists(logo_path):
            run = header_paragraph.add_run()
            run.add_picture(logo_path, width=Inches(1.2))

        # Título
        titulo = doc.add_paragraph()
        titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = titulo.add_run("ACTA DE ENTREGA DE EQUIPOS DE CÓMPUTO")
        run.font.name = 'Calibri'
        run.bold = True
        run.font.size = Pt(14)
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

        doc.add_paragraph("")

        # DATOS DEL COLABORADOR
        tabla = doc.add_table(rows=6, cols=2)
        tabla.style = 'Table Grid'
        tabla.allow_autofit = False

        # Fila 0: Título de la tabla
        tabla.cell(0, 0).merge(tabla.cell(0, 1)).text = "DATOS DEL COLABORADOR"
        celda_titulo = tabla.cell(0, 0).paragraphs[0]
        celda_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_titulo.runs[0] if celda_titulo.runs else celda_titulo.add_run()
        run.bold = True
        sombrear_celda(tabla.cell(0, 0))
        aplicar_fuente_celda(tabla.cell(0, 0))

        # Fila 1 - 5: Llenar la tabla con los datos
        for i, (k, v) in enumerate(datos.items(), start=1):
            tabla.cell(i, 0).text = k
            tabla.cell(i, 1).text = v
            tabla.cell(i, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tabla.cell(i, 1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            sombrear_celda(tabla.cell(i, 0))
            aplicar_fuente_celda(tabla.cell(i, 0))
            aplicar_fuente_celda(tabla.cell(i, 1))

        doc.add_paragraph("")

        # HARDWARE
        tabla_hw = doc.add_table(rows=5, cols=4)
        tabla_hw.style = 'Table Grid'
        tabla_hw.allow_autofit = False

        # Fila 0: Título de la tabla
        tabla_hw.cell(0, 0).merge(tabla_hw.cell(0, 3)).text = "HARDWARE"
        celda_hw = tabla_hw.cell(0, 0).paragraphs[0]
        celda_hw.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_hw.runs[0] if celda_hw.runs else celda_hw.add_run()
        run.bold = True
        sombrear_celda(tabla_hw.cell(0, 0))
        aplicar_fuente_celda(tabla_hw.cell(0, 0))

        # Fila 1 - 2: Encabezados y valores de tipo, marca, modelo, serial
        encabezados1 = ["TIPO", "MARCA", "MODELO", "SERIAL"]
        valores1 = [tipo, marca, modelo, serial]
        for i, texto in enumerate(encabezados1):
            celda = tabla_hw.cell(1, i)
            celda.text = texto
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            for run in celda.paragraphs[0].runs:
                run.bold = True
            tabla_hw.cell(2, i).text = valores1[i]
            aplicar_fuente_celda(tabla_hw.cell(2, i))
        
        # Fila 3 - 4: Encabezados y valores de procesador, RAM, disco, periféricos
        encabezados2 = ["PROCESADOR", "MEMORIA RAM (GB)", "DISCO (GB)", "PERIFERICOS"]
        valores2 = [procesador, ram, disco, perifericos]
        for i, texto in enumerate(encabezados2):
            celda = tabla_hw.cell(3, i)
            celda.text = texto
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            for run in celda.paragraphs[0].runs:
                run.bold = True
            tabla_hw.cell(4, i).text = valores2[i]
            aplicar_fuente_celda(tabla_hw.cell(4, i))

        doc.add_paragraph("")

        # OBSERVACIONES
        tabla_obs = doc.add_table(rows=2, cols=1)
        tabla_obs.style = 'Table Grid'
        tabla_obs.allow_autofit = False

        # Fila 0: Título de la tabla
        celda_encabezado = tabla_obs.cell(0, 0)
        celda_encabezado.text = "OBSERVACIONES"
        sombrear_celda(celda_encabezado)
        aplicar_fuente_celda(celda_encabezado)
        celda_encabezado.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        celda_encabezado.paragraphs[0].runs[0].bold = True

        #Fila 1: Observaciones
        parrafo = tabla_obs.cell(1, 0).add_paragraph(texto_obs)
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        aplicar_fuente_celda(tabla_obs.cell(1, 0))

        doc.add_paragraph("")

        # ENTREGA DE EQUIPO
        tabla_firma = doc.add_table(rows=5, cols=2)
        tabla_firma.style = 'Table Grid'
        tabla_firma.allow_autofit = False

        # Fila 0: Título
        tabla_firma.cell(0, 0).merge(tabla_firma.cell(0, 1)).text = "ENTREGA DE EQUIPO"
        celda_titulo = tabla_firma.cell(0, 0)
        celda_titulo.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        sombrear_celda(celda_titulo)
        aplicar_fuente_celda(celda_titulo)
        celda_titulo.paragraphs[0].runs[0].bold = True

        # Fila 1: Subtítulos "RECIBE" y "ENTREGA"
        tabla_firma.cell(1, 0).text = "RECIBE"
        tabla_firma.cell(1, 1).text = "ENTREGA"
        for col in range(2):
            celda = tabla_firma.cell(1, col)
            celda.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            celda.paragraphs[0].runs[0].bold = True

        # Fila 2: Nombre
        for col, nombre_persona in enumerate([nombre_recibe, nombre_entrega]):
            celda = tabla_firma.cell(2, col)
            p = celda.paragraphs[0]
            run = p.add_run(f"\nNombre: {nombre_persona}\n")
            aplicar_fuente_celda(celda)

        # Fila 3: Firma
        for col in range(2):
            celda = tabla_firma.cell(3, col)
            p = celda.paragraphs[0]
            run = p.add_run("\nFirma:\n")
            aplicar_fuente_celda(celda)

        # Fila 4: Fecha
        for col, fecha in enumerate([fecha_recibe_formateada, fecha_entrega_formateada]):
            celda = tabla_firma.cell(4, col)
            p = celda.paragraphs[0]
            run = p.add_run(f"\nFecha: {fecha}\n")
            aplicar_fuente_celda(celda)

        doc.add_paragraph("")
        doc.add_paragraph("")

        # DATOS DEL EQUIPO
        tabla_equipo = doc.add_table(rows=10, cols=2)
        tabla_equipo.style = 'Table Grid'
        tabla_equipo.allow_autofit = False

        # Fila 0: Título
        tabla_equipo.cell(0, 0).merge(tabla_equipo.cell(0, 1)).text = "DATOS DEL EQUIPO"
        celda_titulo = tabla_equipo.cell(0, 0).paragraphs[0]
        celda_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_titulo.runs[0] if celda_titulo.runs else celda_titulo.add_run()
        run.bold = True
        sombrear_celda(tabla_equipo.cell(0, 0))
        aplicar_fuente_celda(tabla_equipo.cell(0, 0))

        # Fila 1 - 9: Datos del equipo
        for i, (k, v) in enumerate(datos_equipo.items(), start=1):
            tabla_equipo.cell(i, 0).text = k
            tabla_equipo.cell(i, 1).text = v
            sombrear_celda(tabla_equipo.cell(i, 0))
            aplicar_fuente_celda(tabla_equipo.cell(i, 0))
            aplicar_fuente_celda(tabla_equipo.cell(i, 1))

        doc.add_paragraph("")

        # HISTORIAL DE USUARIOS
        tabla_historial = doc.add_table(rows=1, cols=3)
        tabla_historial.style = 'Table Grid'
        tabla_historial.allow_autofit = False

        # Fila 0: Titulo
        tabla_historial.cell(0, 0).merge(tabla_historial.cell(0, 2)).text = "HISTORIAL DE USUARIOS"
        celda_hist_titulo = tabla_historial.cell(0, 0).paragraphs[0]
        celda_hist_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_hist_titulo.runs[0] if celda_hist_titulo.runs else celda_hist_titulo.add_run()
        run.bold = True
        sombrear_celda(tabla_historial.cell(0, 0))
        aplicar_fuente_celda(tabla_historial.cell(0, 0))

        # Fila 1: Encabezados y valores de inicio. fin y usuario
        encabezados_historial = ["INICIO", "FIN", "USUARIO"]
        fila_encabezado = tabla_historial.add_row().cells
        for i, texto in enumerate(encabezados_historial):
            celda = fila_encabezado[i]
            celda.text = texto
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            for run in celda.paragraphs[0].runs:
                run.bold = True

        # Fila 2 - : Datos del historial de usuarios
        filas_agregadas = 0
        for fi, ff, us in zip(historial_inicio, historial_fin, historial_usuario):
            if us.strip():  
                # Agregar formato 'dd-mm-aaaa' en fecha-inicio
                try:
                    fi_formateada = datetime.strptime(fi, "%Y-%m-%d").strftime("%d-%m-%Y")
                except (ValueError, TypeError):
                    fi_formateada = fi
                # Agregar formato 'dd-mm-aaaa' en fecha-fin
                if ff == "ACTUAL":
                    ff_formateada = "ACTUAL"
                else:
                    try:
                        ff_formateada = datetime.strptime(ff, "%Y-%m-%d").strftime("%d-%m-%Y")
                    except (ValueError, TypeError):
                        ff_formateada = ff

                fila = tabla_historial.add_row().cells
                fila[0].text = fi_formateada
                fila[1].text = ff_formateada
                fila[2].text = us
                for celda in fila:
                    aplicar_fuente_celda(celda)
                filas_agregadas += 1

        # Agrega filas vacías a la tabla de historial para asegurar que tenga 8 filas en total
        total_filas_deseadas = 8  
        filas_existentes = filas_agregadas + 2  
        filas_faltantes = total_filas_deseadas - filas_existentes
        for _ in range(filas_faltantes):
            fila = tabla_historial.add_row().cells
            for celda in fila:
                celda.text = ""
                aplicar_fuente_celda(celda)

        doc.add_paragraph("")

        # HISTORIAL DE EVENTOS
        tabla_eventos = doc.add_table(rows=1, cols=2)
        tabla_eventos.style = 'Table Grid'
        tabla_eventos.allow_autofit = False

        # Fila 0: Título de la tabla
        tabla_eventos.cell(0, 0).merge(tabla_eventos.cell(0, 1)).text = "HISTORIAL DE EVENTOS"
        celda_eventos_titulo = tabla_eventos.cell(0, 0).paragraphs[0]
        celda_eventos_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_eventos_titulo.runs[0] if celda_eventos_titulo.runs else celda_eventos_titulo.add_run()
        run.bold = True
        sombrear_celda(tabla_eventos.cell(0, 0))
        aplicar_fuente_celda(tabla_eventos.cell(0, 0))

        # Fila 1: Encabezados de fecha y observaciones
        encabezados_eventos = ["FECHA", "OBSERVACIONES"]
        fila_encabezado = tabla_eventos.add_row().cells
        for i, texto in enumerate(encabezados_eventos):
            celda = fila_encabezado[i]
            celda.text = texto
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            for run in celda.paragraphs[0].runs:
                run.bold = True

        # Fila 2 - : Datos del historial de eventos
        filas_evento_agregadas = 0
        for fecha, obs in zip(evento_fechas, evento_observaciones):
            if obs.strip():  
                try:
                    fecha_formateada = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d-%m-%Y")
                except (ValueError, TypeError):
                    fecha_formateada = fecha

                fila = tabla_eventos.add_row().cells
                fila[0].text = fecha_formateada
                fila[1].text = obs
                for celda in fila:
                    aplicar_fuente_celda(celda)
                filas_evento_agregadas += 1

        # Añade filas vacías a la tabla de eventos hasta completar un total de 8 filas visibles
        total_filas_eventos_deseadas = 8 
        filas_eventos_existentes = filas_evento_agregadas + 2 
        filas_eventos_faltantes = total_filas_eventos_deseadas - filas_eventos_existentes
        for _ in range(filas_eventos_faltantes):
            fila = tabla_eventos.add_row().cells
            for celda in fila:
                celda.text = ""
                aplicar_fuente_celda(celda)

        doc.add_paragraph("")

        # MANTENIMIENTO DE HARDWARE
        tabla_hardware = doc.add_table(rows=17, cols=3)
        tabla_hardware.style = 'Table Grid'
        tabla_hardware.allow_autofit = False

        # Ancho de columnas
        anchos = [8, 2.25, 5] # primera columna más ancha
        for col_idx, ancho in enumerate(anchos):
            for row_idx in range(len(tabla_hardware.rows)):
                tabla_hardware.cell(row_idx, col_idx).width = Cm(ancho)

        # MANTENIMIENTO DE HARDWARE
        tabla_hardware.cell(0, 0).merge(tabla_hardware.cell(0, 2)).text = "MANTENIMIENTO DE HARDWARE"
        celda_titulo = tabla_hardware.cell(0, 0).paragraphs[0]
        celda_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_titulo.runs[0] if celda_titulo.runs else celda_titulo.add_run()
        run.bold = True
        sombrear_celda(tabla_hardware.cell(0, 0))
        aplicar_fuente_celda(tabla_hardware.cell(0, 0))

        # Encabezados
        tabla_hardware.cell(1, 0).text = "ACTIVIDAD"
        tabla_hardware.cell(1, 1).text = "SI / NO"
        tabla_hardware.cell(1, 2).text = "DETALLES"

        # Aplica sombreado, fuente y alineación centrada a las 3 primeras columnas del encabezado.
        for col in range(3):
            celda = tabla_hardware.cell(1, col)
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            celda.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = celda.paragraphs[0].runs[0]
            run.bold = True

        # Recorre cada par (pregunta, clave)
        for i, (texto_pregunta, clave) in enumerate(preguntas, start=2):
            tabla_hardware.cell(i, 0).text = texto_pregunta
            sombrear_celda(tabla_hardware.cell(i, 0))
            aplicar_fuente_celda(tabla_hardware.cell(i, 0))

            valor_sn = request.form.get(f"{clave}_sn", "false")
            texto_sn = "Sí" if valor_sn.lower() == "true" else "No"
            tabla_hardware.cell(i, 1).text = texto_sn
            aplicar_fuente_celda(tabla_hardware.cell(i, 1))

            detalle = request.form.get(f"{clave}_detalle", "")
            tabla_hardware.cell(i, 2).text = detalle
            aplicar_fuente_celda(tabla_hardware.cell(i, 2))

        doc.add_paragraph("")

        # MANTENIMIENTO DE SOFTWARE
        tabla_software = doc.add_table(rows=10, cols=3)
        tabla_software.style = 'Table Grid'
        tabla_software.allow_autofit = False

        # Ajustar anchos de columna
        anchos = [8, 2.25, 5]
        for col_idx, ancho in enumerate(anchos):
            for row_idx in range(len(tabla_software.rows)):
                tabla_software.cell(row_idx, col_idx).width = Cm(ancho)

        # Título: combinar las 3 columnas en la primera fila
        tabla_software.cell(0, 0).merge(tabla_software.cell(0, 2)).text = "MANTENIMIENTO DE SOFTWARE"
        celda_titulo = tabla_software.cell(0, 0).paragraphs[0]
        celda_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = celda_titulo.runs[0] if celda_titulo.runs else celda_titulo.add_run()
        run.bold = True
        sombrear_celda(tabla_software.cell(0, 0))
        aplicar_fuente_celda(tabla_software.cell(0, 0))

        # Encabezados
        tabla_software.cell(1, 0).text = "ACTIVIDAD"
        tabla_software.cell(1, 1).text = "SI / NO"
        tabla_software.cell(1, 2).text = "DETALLES"

        # Aplica sombreado, fuente y alineación centrada a las 3 primeras columnas del encabezado.
        for col in range(3):
            celda = tabla_software.cell(1, col)
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)
            celda.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = celda.paragraphs[0].runs[0]
            run.bold = True

        # Recorre cada par (pregunta, clave)
        for i, (pregunta, clave) in enumerate(preguntas_software, start=2):
            tabla_software.cell(i, 0).text = pregunta
            sombrear_celda(tabla_software.cell(i, 0))
            aplicar_fuente_celda(tabla_software.cell(i, 0))

            valor_sn = request.form.get(f"{clave}_sn", "false")
            texto_sn = "Sí" if valor_sn.lower() == "true" else "No"
            tabla_software.cell(i, 1).text = texto_sn
            aplicar_fuente_celda(tabla_software.cell(i, 1))

            detalle = request.form.get(f"{clave}_detalle", "")
            tabla_software.cell(i, 2).text = detalle
            aplicar_fuente_celda(tabla_software.cell(i, 2))

        doc.add_paragraph()

        # LISTA DE PROGRAMAS POR AREA
        titulo = doc.add_paragraph()
        titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = titulo.add_run("LISTA DE PROGRAMAS POR ÁREA")
        run.bold = True
        run.font.size = Pt(12)

        # Aplica sombreado, fuente y alineación
        run.font.name = 'Calibri'
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        tabla_obs = doc.add_table(rows=9, cols=5)
        tabla_obs.style = 'Table Grid'
        tabla_obs.allow_autofit = False

        # Encabezado Especifico
        celda_esp = tabla_obs.cell(0, 1).merge(tabla_obs.cell(0, 4))
        celda_esp.text = "Específico"
        parrafo = celda_esp.paragraphs[0]
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        sombrear_celda(celda_esp)
        aplicar_fuente_celda(celda_esp)

        # Encabezado General (Todos)
        celda_general = tabla_obs.cell(1, 0).merge(tabla_obs.cell(0, 0))
        celda_general.text = "General (Todos)"
        parrafo = celda_general.paragraphs[0]
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        celda_general.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        sombrear_celda(celda_general)
        aplicar_fuente_celda(celda_general)

        # Lista de áreas para insertar
        areas = ["AUDIT", "AOS", "ADMIN", "TAX & LEGAL"]

        for col_idx, area in enumerate(areas, start=1):
            celda = tabla_obs.cell(1, col_idx)
            celda.text = area

            # Centrado horizontal
            parrafo = celda.paragraphs[0]
            parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
            run.bold = True
            run.font.name = "Calibri"
            run.font.size = Pt(11)

            # Sombreado y fuente
            sombrear_celda(celda)
            aplicar_fuente_celda(celda)

        # Iteramos por columnas y filas, llenando la tabla
        for col_idx, programas in enumerate(programas_por_columna):
            for row_offset, programa in enumerate(programas, start=2):  # filas 2 a 8
                celda = tabla_obs.cell(row_offset, col_idx)
                celda.text = programa
                parrafo = celda.paragraphs[0]
                # Aseguramos que haya un run para configurar la fuente
                run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
                run.font.name = "Calibri"
                run.font.size = Pt(11)

        # Guardamos el docx en memoria
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)  # Volver al inicio del archivo en memoria

        # Enviar el archivo para descarga
        return send_file(
            file_stream,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    return render_template('formulario.html', fecha_actual=fecha_actual)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)




