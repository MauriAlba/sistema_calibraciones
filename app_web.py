# from flask import Flask, render_template, request, redirect
# from flask import send_file
from flask import Flask, render_template, request, redirect, send_file, session
from openpyxl.styles import Font, PatternFill, Alignment
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
import sqlite3
import datetime
import os


# USUARIO = "admin"
# PASSWORD = "1234"

app = Flask(__name__)


app.secret_key = "clave_secreta"


# def crear_base():
#     conexion = sqlite3.connect("calibraciones_4.db")
#     cursor = conexion.cursor()

#     cursor.execute("""
#     CREATE TABLE IF NOT EXISTS calibraciones_4 (
#         id INTEGER PRIMARY KEY AUTOINCREMENT,
#         instrumento TEXT,
#         tipo TEXT,
#         fecha TEXT,
#         frecuencia INTEGER
#     )
#     """)

#     conexion.commit()
#     conexion.close()

# crear_base()

def crear_base():
    conexion = sqlite3.connect("calibraciones_4.db")
    cursor = conexion.cursor()

    # Tabla calibraciones
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS calibraciones_4 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        instrumento TEXT,
        tipo TEXT,
        fecha TEXT,
        frecuencia INTEGER
    )
    """)

    # 👉 Nueva tabla usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE,
        password TEXT
    )
    """)

    conexion.commit()
    conexion.close()

# def crear_usuario_inicial():
#     conexion = sqlite3.connect("calibraciones_4.db")
#     cursor = conexion.cursor()

#     try:
#         cursor.execute(
#             "INSERT INTO usuarios (usuario, password) VALUES (?, ?)",
#             ("oper", "6789")
#         )
#         conexion.commit()
#     except:
#         pass  # si ya existe, no hace nada

#     conexion.close()

def crear_usuario_inicial():
    conexion = sqlite3.connect("calibraciones_4.db")
    cursor = conexion.cursor()

    try:
        password_hash = generate_password_hash("1234")

        cursor.execute(
            "INSERT INTO usuarios (usuario, password) VALUES (?, ?)",
            ("admin", password_hash)
        )

        conexion.commit()
    except:
        pass

    conexion.close()

crear_base()
crear_usuario_inicial()

@app.route("/")
def index():
    if not session.get("logueado"):
        return redirect("/login")

    
    conexion = sqlite3.connect("calibraciones_4.db")
    cursor = conexion.cursor()

    cursor.execute("SELECT instrumento, tipo, fecha, frecuencia FROM calibraciones_4")
    datos_db = cursor.fetchall()

    conexion.close()

    datos = []
    hoy = datetime.date.today()

    for instrumento, tipo, fecha, frecuencia in datos_db:
        fecha_cal = datetime.datetime.strptime(fecha, "%Y-%m-%d").date()
        proxima = fecha_cal + datetime.timedelta(days=frecuencia)

        dias_restantes = (proxima - hoy).days

        if hoy >= proxima:
            estado = "vencido"
        elif dias_restantes <= 7:
            estado = "proximo"
        else:
            estado = "ok"

        datos.append({
            "instrumento": instrumento,
            "tipo": tipo,
            "fecha": fecha,
            "frecuencia": frecuencia,
            "proxima": proxima,
            "estado": estado
        })

    return render_template("index.html", datos=datos)


@app.route("/agregar", methods=["POST"])
def agregar():
    if not session.get("logueado"):
        return redirect("/login")
    

    instrumento = request.form["instrumento"]
    tipo = request.form["tipo"]
    frecuencia = request.form["frecuencia"]

    fecha = datetime.date.today()

    conexion = sqlite3.connect("calibraciones_4.db")
    cursor = conexion.cursor()

    cursor.execute(
        "INSERT INTO calibraciones_4 (instrumento, tipo, fecha, frecuencia) VALUES (?, ?, ?, ?)",
        (instrumento, tipo, fecha, frecuencia)
    )

    conexion.commit()
    conexion.close()

    return redirect("/")

@app.route("/exportar")
def exportar():
    if not session.get("logueado"):
        return redirect("/login")
    
    conexion = sqlite3.connect("calibraciones_4.db")
    cursor = conexion.cursor()

    # cursor.execute("SELECT instrumento, tipo, fecha, frecuencia FROM calibraciones_4")
    # datos_db = cursor.fetchall()
    # conexion.close()

    # cursor.execute("""
    # SELECT instrumento, tipo, fecha, frecuencia
    # FROM calibraciones_4
    # GROUP BY instrumento
    # ORDER BY instrumento
    # """)
    # datos_db = cursor.fetchall()
    # conexion.close()

    # cursor.execute("""
    # SELECT instrumento, tipo, fecha, frecuencia
    # FROM calibraciones_4
    # ORDER BY instrumento
    # """)
    cursor.execute("""
    SELECT instrumento, tipo, fecha, frecuencia
    FROM calibraciones_4
    """)

    datos_db = cursor.fetchall()
    conexion.close()

    # cursor.execute("""
    # SELECT instrumento, tipo, fecha, frecuencia
    # FROM calibraciones_4
    # ORDER BY instrumento, fecha DESC
    # """)

    # cursor.execute("""
    # SELECT c1.instrumento, c1.tipo, c1.fecha, c1.frecuencia
    # FROM calibraciones_4 c1
    # INNER JOIN (
    #     SELECT instrumento, MAX(fecha) as max_fecha
    #     FROM calibraciones_4
    #     GROUP BY instrumento
    # ) c2
    # ON c1.instrumento = c2.instrumento AND c1.fecha = c2.max_fecha
    # ORDER BY c1.instrumento
    # """)
    # datos_db = cursor.fetchall()
    # conexion.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calibraciones"

    # Encabezados
    #ws.append(["Instrumento", "Tipo", "Última", "Próxima", "Estado"])
    headers = ["Instrumento", "Tipo", "Última", "Próxima", "Estado"]
    ws.append(headers)

    for col in ws[1]:
        col.font = Font(bold=True)
        col.alignment = Alignment(horizontal="center")

    hoy = datetime.date.today()

    #for instrumento, tipo, fecha, frecuencia in datos_db:
    #    fecha_cal = datetime.datetime.strptime(fecha, "%Y-%m-%d").date()
    #    proxima = fecha_cal + datetime.timedelta(days=frecuencia)

    #    if hoy >= proxima:
    #        estado = "VENCIDO"
    #    elif (proxima - hoy).days <= 7:
    #        estado = "PRÓXIMO"
    #    else:
    #        estado = "OK"

    #    ws.append([instrumento, tipo, str(fecha_cal), str(proxima), estado])
    for instrumento, tipo, fecha, frecuencia in datos_db:
        fecha_cal = datetime.datetime.strptime(fecha, "%Y-%m-%d").date()
        proxima = fecha_cal + datetime.timedelta(days=frecuencia)

        if hoy >= proxima:
            estado = "VENCIDO"
            color = "FFCCCC"  # rojo
        elif (proxima - hoy).days <= 7:
            estado = "PRÓXIMO"
            color = "FFF3CD"  # amarillo
        else:
            estado = "OK"
            color = "CCFFCC"  # verde

        fila = [instrumento, tipo, str(fecha_cal), str(proxima), estado]
        ws.append(fila)

        fila_excel = ws.max_row

        for col in range(1, 6):
            ws.cell(row=fila_excel, column=col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    columnas = ["A", "B", "C", "D", "E"]

    for col in columnas:
        ws.column_dimensions[col].width = 20

        # archivo = "reporte_calibraciones.xlsx"
        # wb.save(archivo)

        # return send_file(archivo, as_attachment=True)
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
    archivo = "reporte_calibraciones.xlsx"
    wb.save(archivo)

    return send_file(archivo, as_attachment=True)


# @app.route("/login", methods=["GET", "POST"])
# def login():
#     if request.method == "POST":
#         user = request.form["usuario"]
#         password = request.form["password"]

#         conexion = sqlite3.connect("calibraciones_4.db")
#         cursor = conexion.cursor()

#         cursor.execute(
#             "SELECT * FROM usuarios WHERE usuario=? AND password=?",
#             (user, password)
#         )

#         resultado = cursor.fetchone()
#         conexion.close()

#         if resultado:
#             session["logueado"] = True
#             session["usuario"] = user
#             return redirect("/")
#         else:
#             return "❌ Usuario o contraseña incorrectos"

#     return render_template("login.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form["usuario"]
        password = request.form["password"]

        conexion = sqlite3.connect("calibraciones_4.db")
        cursor = conexion.cursor()

        cursor.execute(
            "SELECT password FROM usuarios WHERE usuario=?",
            (user,)
        )

        resultado = cursor.fetchone()
        conexion.close()

        if resultado:
            password_hash = resultado[0]

            if check_password_hash(password_hash, password):
                session["logueado"] = True
                session["usuario"] = user
                return redirect("/")
        
        return "❌ Usuario o contraseña incorrectos"

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


#app.run(debug=True)
#app.run(host="0.0.0.0", port=5001, debug=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))