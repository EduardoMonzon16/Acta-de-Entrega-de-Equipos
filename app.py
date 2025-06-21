# -*- coding: utf-8 -*-
"""
Sistema de Generación de Actas de Entrega de Equipos de Cómputo
Forvis Mazars Perú
"""

# ============================================================================
# IMPORTS
# ============================================================================
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

# ============================================================================
# CONFIGURACIÓN DE LA APLICACIÓN
# ============================================================================
app = Flask(__name__)
app.secret_key = 'clave_secreta_segura'

# ============================================================================
# CONFIGURACIÓN DE BASE DE DATOS
# ============================================================================
def conectar():
    """Establece conexión con la base de datos MySQL"""
    return mysql.connector.connect(
        host='localhost',          
        user='root',   
        password='Universitario12#',
        database='systembd',
        port=3307 
    )

# ============================================================================
# FUNCIONES DE AUTENTICACIÓN
# ============================================================================
def validar_credenciales(usuario, password):
    """Valida las credenciales del usuario y devuelve su ID si son correctas"""
    try:
        conexion = conectar()
        cursor = conexion.cursor()
        sql = "SELECT idUsuarios FROM usuarios WHERE Nombre = %s AND Contraseña = %s"
        cursor.execute(sql, (usuario, password))
        resultado = cursor.fetchone()
        cursor.close()
        conexion.close()

        if resultado:
            return resultado[0]  # idUsuarios
        else:
            return None
    except Error as e:
        print(f"Error en la conexión: {e}")
        return None

# ============================================================================
# FUNCIONES AUXILIARES PARA DOCX
# ============================================================================
def sombrear_celda(celda, color_hex="D9D9D9"):
    """Aplica color de fondo a una celda"""
    tc = celda._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)
    
def aplicar_fuente_celda(celda, fuente="Calibri", tam=11):
    """Establece la fuente y el tamaño del texto en una celda"""
    for parrafo in celda.paragraphs:
        for run in parrafo.runs:
            run.font.name = fuente
            run.font.size = Pt(tam)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), fuente)

def formatear_fecha(fecha_str):
    """Convierte fecha de formato YYYY-MM-DD a DD-MM-YYYY"""
    try:
        return datetime.strptime(fecha_str, "%Y-%m-%d").strftime("%d-%m-%Y")
    except (ValueError, TypeError):
        return fecha_str

# ============================================================================
# DATOS ESTÁTICOS
# ============================================================================
def obtener_preguntas_hardware():
    """Retorna las preguntas para mantenimiento de hardware"""
    return [
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

def obtener_preguntas_software():
    """Retorna las preguntas para mantenimiento de software"""
    return [
        ("¿Se verificaron los programas vigentes?", "programas_vigentes"),
        ("¿Se eliminaron los programas externos?", "programas_externos"),
        ("¿Se realizó la limpieza de archivos temporales?", "limpieza_temporal"),
        ("¿Se eliminaron los perfiles antiguos?", "perfiles_antiguos"),
        ("¿Se realizaron actualizaciones de Windows?", "actualizaciones_windows"),
        ("¿Se comprobó el estado del disco duro?", "estado_disco"),
        ("¿Se realizo el backup del usuario anterior?", "backup_usuario"),
        ("¿Se realizó desfragmentación del disco duro?", "desfragmentacion"),
    ]

def obtener_programas_por_area():
    """Retorna la lista de programas organizados por área"""
    return [
        ["Anydesk", "Office 365", "Cisco VPN", "PDF24", "Microsoft Teams", "Microsoft Defender", "Acrobat Reader"],
        ["Atlas", "Auditsoft", "---", "---", "---", "---", "---"],
        ["Concar", "Starsoft", "PDT", "PLAME", "---", "---", "---"],
        ["Impresoras", "Scanners", "---", "---", "---", "---", "---"],
        ["PDT", "PLAME", "PLE", "Renta Anual", "Mis declaraciones", "PDB", "---"]
    ]

def obtener_texto_observaciones():
    """Retorna el texto estándar de observaciones"""
    return (
        "Certifico que los elementos detallados en el presente documento me han sido entregados para mi cuidado "
        "y custodia con el propósito de cumplir con las tareas y asignaciones propias de mi cargo, siendo estos "
        "de mi única y exclusiva responsabilidad. Me comprometo a usar correctamente los recursos solo para los "
        "fines establecidos, y a no instalar ni permitir la instalación de software por personal ajeno al personal "
        "de TI de Forvis Mazars Perú. De igual forma me comprometo a devolver el equipo en las mismas condiciones "
        "y con los mismos accesorios que me fue entregado, cuando se me programe algún cambio de equipo o el "
        "vínculo laboral haya culminado."
    )

# ============================================================================
# FUNCIONES DE PROCESAMIENTO DE DATOS
# ============================================================================
def obtener_cargos():
    """Obtiene todos los cargos disponibles desde la base de datos"""
    try:
        conexion = conectar()
        cursor = conexion.cursor()
        sql = "SELECT idCargos, NombreCargo FROM cargos ORDER BY idCargos"
        cursor.execute(sql)
        resultados = cursor.fetchall()
        cursor.close()
        conexion.close()
        
        # Retorna una lista de tuplas (id, nombre)
        return resultados
    except Error as e:
        print(f"Error al obtener cargos: {e}")
        return []

def obtener_cargo_por_id(id_cargo):
    """Obtiene un cargo específico por su ID"""
    try:
        conexion = conectar()
        cursor = conexion.cursor()
        sql = "SELECT NombreCargo FROM cargos WHERE idCargos = %s"
        cursor.execute(sql, (id_cargo,))
        resultado = cursor.fetchone()
        cursor.close()
        conexion.close()
        
        return resultado[0] if resultado else None
    except Error as e:
        print(f"Error al obtener cargo por ID: {e}")
        return None

def procesar_datos_formulario(request):
    """Procesa y organiza todos los datos del formulario"""
    # Datos personales
    datos_personales = {
        "nombre": request.form['nombre'],
        "correo": f"{request.form['correo'].strip()}@forvismazars.com",
        "cargo": request.form['cargo'],
        "usuario": request.form['usuario'],
        "telefono": request.form['telefono']
    }
    
    # Hardware
    datos_hardware = {
        "tipo": request.form['tipo'],
        "marca": request.form['marca'],
        "modelo": request.form['modelo'],
        "serial": request.form['serial'],
        "procesador": request.form['procesador'],
        "ram": request.form['ram'],
        "disco": request.form['disco'],
        "perifericos": request.form['perifericos']
    }
    
    # Datos del equipo
    datos_equipo = {
        "fecha_compra": formatear_fecha(request.form.get('fecha_compra', '')),
        "equipo": request.form.get('equipo', ''),
        "marca_equipo": request.form.get('marca_equipo', ''),
        "hostname": request.form.get('hostname', ''),
        "modelo_equipo": request.form.get('modelo_equipo', ''),
        "detalle": request.form.get('detalle', ''),
        "serie_equipo": request.form.get('serie_equipo', ''),
        "os_equipo": request.form.get('os', ''),
        "garantia": formatear_fecha(request.form.get('garantia', ''))
    }
    
    # Entrega de equipo
    datos_entrega = {
        "nombre_recibe": request.form.get('nombre_recibe', ''),
        "fecha_recibe": formatear_fecha(request.form.get('fecha_recibe', '')),
        "nombre_entrega": request.form.get('nombre_entrega', ''),
        "fecha_entrega": formatear_fecha(request.form.get('fecha_entrega', ''))
    }
    
    # Historial de usuarios
    historial_usuarios = list(zip(
        request.form.getlist('historial_inicio[]'),
        request.form.getlist('historial_fin[]'),
        request.form.getlist('historial_usuario[]')
    ))
    
    # Historial de eventos
    historial_eventos = list(zip(
        request.form.getlist('evento_fecha[]'),
        request.form.getlist('evento_observaciones[]')
    ))
    
    return {
        'personales': datos_personales,
        'hardware': datos_hardware,
        'equipo': datos_equipo,
        'entrega': datos_entrega,
        'historial_usuarios': historial_usuarios,
        'historial_eventos': historial_eventos
    }

# ============================================================================
# FUNCIONES DE GENERACIÓN DE DOCUMENTO
# ============================================================================
def crear_encabezado(doc):
    """Crea el encabezado del documento con el logo"""
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

def crear_titulo(doc):
    """Crea el título principal del documento"""
    titulo = doc.add_paragraph()
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = titulo.add_run("ACTA DE ENTREGA DE EQUIPOS DE CÓMPUTO")
    run.font.name = 'Calibri'
    run.bold = True
    run.font.size = Pt(14)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
    doc.add_paragraph("")

def crear_tabla_datos_colaborador(doc, datos_personales):
    """Crea la tabla con los datos del colaborador"""
    datos = {
        "Nombre:": datos_personales['nombre'],
        "Correo:": datos_personales['correo'],
        "Cargo:": datos_personales['cargo'],
        "Usuario de red:": datos_personales['usuario'],
        "Teléfono:": datos_personales['telefono']
    }
    
    tabla = doc.add_table(rows=6, cols=2)
    tabla.style = 'Table Grid'
    tabla.allow_autofit = False

    # Título
    tabla.cell(0, 0).merge(tabla.cell(0, 1)).text = "DATOS DEL COLABORADOR"
    celda_titulo = tabla.cell(0, 0).paragraphs[0]
    celda_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = celda_titulo.runs[0] if celda_titulo.runs else celda_titulo.add_run()
    run.bold = True
    sombrear_celda(tabla.cell(0, 0))
    aplicar_fuente_celda(tabla.cell(0, 0))

    # Datos
    for i, (k, v) in enumerate(datos.items(), start=1):
        tabla.cell(i, 0).text = k
        tabla.cell(i, 1).text = v
        tabla.cell(i, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        tabla.cell(i, 1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        sombrear_celda(tabla.cell(i, 0))
        aplicar_fuente_celda(tabla.cell(i, 0))
        aplicar_fuente_celda(tabla.cell(i, 1))

    doc.add_paragraph("")

def crear_tabla_hardware(doc, datos_hardware):
    """Crea la tabla con información del hardware"""
    tabla_hw = doc.add_table(rows=5, cols=4)
    tabla_hw.style = 'Table Grid'
    tabla_hw.allow_autofit = False

    # Título
    tabla_hw.cell(0, 0).merge(tabla_hw.cell(0, 3)).text = "HARDWARE"
    celda_hw = tabla_hw.cell(0, 0).paragraphs[0]
    celda_hw.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = celda_hw.runs[0] if celda_hw.runs else celda_hw.add_run()
    run.bold = True
    sombrear_celda(tabla_hw.cell(0, 0))
    aplicar_fuente_celda(tabla_hw.cell(0, 0))

    # Primera fila de datos
    encabezados1 = ["TIPO", "MARCA", "MODELO", "SERIAL"]
    valores1 = [datos_hardware['tipo'], datos_hardware['marca'], 
                datos_hardware['modelo'], datos_hardware['serial']]
    
    for i, texto in enumerate(encabezados1):
        celda = tabla_hw.cell(1, i)
        celda.text = texto
        sombrear_celda(celda)
        aplicar_fuente_celda(celda)
        for run in celda.paragraphs[0].runs:
            run.bold = True
        tabla_hw.cell(2, i).text = valores1[i]
        aplicar_fuente_celda(tabla_hw.cell(2, i))
    
    # Segunda fila de datos
    encabezados2 = ["PROCESADOR", "MEMORIA RAM (GB)", "DISCO (GB)", "PERIFERICOS"]
    valores2 = [datos_hardware['procesador'], datos_hardware['ram'], 
                datos_hardware['disco'], datos_hardware['perifericos']]
    
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

def crear_tabla_observaciones(doc):
    """Crea la tabla de observaciones"""
    tabla_obs = doc.add_table(rows=2, cols=1)
    tabla_obs.style = 'Table Grid'
    tabla_obs.allow_autofit = False

    # Título
    celda_encabezado = tabla_obs.cell(0, 0)
    celda_encabezado.text = "OBSERVACIONES"
    sombrear_celda(celda_encabezado)
    aplicar_fuente_celda(celda_encabezado)
    celda_encabezado.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    celda_encabezado.paragraphs[0].runs[0].bold = True

    # Contenido
    parrafo = tabla_obs.cell(1, 0).add_paragraph(obtener_texto_observaciones())
    parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    aplicar_fuente_celda(tabla_obs.cell(1, 0))

    doc.add_paragraph("")

def crear_tabla_entrega(doc, datos_entrega):
    """Crea la tabla de entrega de equipo"""
    tabla_firma = doc.add_table(rows=5, cols=2)
    tabla_firma.style = 'Table Grid'
    tabla_firma.allow_autofit = False

    # Título
    tabla_firma.cell(0, 0).merge(tabla_firma.cell(0, 1)).text = "ENTREGA DE EQUIPO"
    celda_titulo = tabla_firma.cell(0, 0)
    celda_titulo.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    sombrear_celda(celda_titulo)
    aplicar_fuente_celda(celda_titulo)
    celda_titulo.paragraphs[0].runs[0].bold = True

    # Subtítulos
    tabla_firma.cell(1, 0).text = "RECIBE"
    tabla_firma.cell(1, 1).text = "ENTREGA"
    for col in range(2):
        celda = tabla_firma.cell(1, col)
        celda.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        sombrear_celda(celda)
        aplicar_fuente_celda(celda)
        celda.paragraphs[0].runs[0].bold = True

    # Datos
    nombres = [datos_entrega['nombre_recibe'], datos_entrega['nombre_entrega']]
    fechas = [datos_entrega['fecha_recibe'], datos_entrega['fecha_entrega']]
    
    # Nombres
    for col, nombre_persona in enumerate(nombres):
        celda = tabla_firma.cell(2, col)
        p = celda.paragraphs[0]
        run = p.add_run(f"\nNombre: {nombre_persona}\n")
        aplicar_fuente_celda(celda)

    # Firmas
    for col in range(2):
        celda = tabla_firma.cell(3, col)
        p = celda.paragraphs[0]
        run = p.add_run("\nFirma:\n")
        aplicar_fuente_celda(celda)

    # Fechas
    for col, fecha in enumerate(fechas):
        celda = tabla_firma.cell(4, col)
        p = celda.paragraphs[0]
        run = p.add_run(f"\nFecha: {fecha}\n")
        aplicar_fuente_celda(celda)

    doc.add_paragraph("")
    doc.add_paragraph("")

# ============================================================================
# RUTAS DE LA APLICACIÓN
# ============================================================================
@app.route('/')
def index():
    """Redirige la página principal al login"""
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Maneja el login de usuarios"""
    if request.method == 'POST':
        usuario = request.form['usuario']
        password = request.form['password']
        
        id_usuario = validar_credenciales(usuario, password)
        
        if id_usuario:
            session['usuario'] = usuario
            session['id_usuario'] = id_usuario
            
            # Redirigir según el ID específico
            if id_usuario == 1:
                return redirect(url_for('administradorti'))
            elif id_usuario == 2:
                return redirect(url_for('usuarioti'))
            else:
                # Si no es 1 ni 2, mostrar mensaje y limpiar sesión
                session.clear()
                flash('No tienes permisos para acceder al sistema', 'error')
        else:
            flash('Usuario o contraseña incorrectos', 'error')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Cierra la sesión del usuario"""
    session.clear()
    return redirect(url_for('login'))

@app.route('/administradorti')
def administradorti():
    """Ruta para administrador - solo ID 1"""
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    if session.get('id_usuario') != 1:
        flash('No tienes permisos para acceder a esta sección', 'error')
        return redirect(url_for('login'))
    
    return render_template('administradorti.html')

@app.route('/usuarioti', methods=['GET', 'POST'])
def usuarioti():
    """Ruta principal que maneja el formulario y genera el documento - solo ID 2"""
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    if session.get('id_usuario') != 2:
        flash('No tienes permisos para acceder a esta sección', 'error')
        return redirect(url_for('login'))
    
    fecha_actual = datetime.today().strftime('%Y-%m-%d')
    cargos = obtener_cargos()
    
    if request.method == 'POST':
        # Procesar datos del formulario
        datos = procesar_datos_formulario(request)
        
        # Crear nombre único para el documento
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"Acta_Entrega_{datos['personales']['nombre'].replace(' ', '_')}_{timestamp}.docx"
        
        # Crear documento
        doc = Document()
        
        # Generar secciones del documento
        crear_encabezado(doc)
        crear_titulo(doc)
        crear_tabla_datos_colaborador(doc, datos['personales'])
        crear_tabla_hardware(doc, datos['hardware'])
        crear_tabla_observaciones(doc)
        crear_tabla_entrega(doc, datos['entrega'])
        
        # TODO: Agregar las demás tablas (datos del equipo, historial, mantenimiento, etc.)
        # Esta es una versión simplificada para mostrar la organización
        
        # Guardar documento en memoria
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        # Enviar archivo para descarga
        return send_file(
            file_stream,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    return render_template('usuarioti.html', fecha_actual=fecha_actual, cargos=cargos)

# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)