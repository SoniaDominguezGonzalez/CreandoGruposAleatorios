import pandas as pd
import itertools
import random
from fpdf import FPDF
## IMPORTANTE
# Comprobar que los nombres tanto del archivo .xlsx como de las columnas coinciden con los indicados en el código!

# === CONFIG ===
RUTA_EXCEL = 'Listado_alumnado.xlsx'  # Cambiar por el archivo que contenga el listado del alumnado
COL_NOMBRE = 'Nombre'  # Cambiar por el nombre de la columna que contenga el nombre del alumnado
COL_APELLIDO1 = 'Primer apellido'  # Cambiar por el nombre de la columna que contenga el primer apellido del alumnado
COL_APELLIDO2 = 'Segundo apellido'  # Cambiar por el nombre de la columna que contenga el segundo apellido del alumnado
INTEGRANTES_POR_GRUPO = 3
SESIONES_MAXIMAS = 16
INTENTOS_POR_SESION = 1000 # No cambiar

def obtener_parejas(grupo):
    return set(itertools.combinations(sorted(grupo), 2))

def generar_sesiones(alumnos, num_sesiones, tam_grupo, max_intentos=1000):
    parejas_usadas = set()
    sesiones = []
    alumnos_idx = list(range(len(alumnos)))
    grupos_posibles = list(itertools.combinations(alumnos_idx, tam_grupo))
    for sesion_num in range(num_sesiones):
        exito = False
        for intento in range(max_intentos):
            random.shuffle(grupos_posibles)
            sesion = []
            alumnos_en_sesion = set()
            parejas_en_sesion = set()
            for grupo in grupos_posibles:
                if any(a in alumnos_en_sesion for a in grupo):
                    continue
                parejas_grupo = obtener_parejas(grupo)
                if parejas_grupo & parejas_usadas:
                    continue
                sesion.append(grupo)
                alumnos_en_sesion.update(grupo)
                parejas_en_sesion.update(parejas_grupo)
                if len(sesion) == len(alumnos) // tam_grupo:
                    break
            if len(sesion) == len(alumnos) // tam_grupo:
                sesiones.append(sesion)
                parejas_usadas.update(parejas_en_sesion)
                exito = True
                print(f"Sesión {sesion_num+1} generada en intento {intento+1}")
                break
        if not exito:
            print(f"No se pudo generar la sesión {sesion_num+1} tras {max_intentos} intentos.")
            break
    return sesiones

def exportar_pdf(sesiones, alumnos, tam_grupo, nombre_archivo='listado_grupos_sesiones.pdf'):  #Cambiar el nombre del PDF si lo desea
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for i, sesion in enumerate(sesiones, 1):
        if i > 1:
            pdf.add_page()  # Nueva página para cada sesión excepto la primera
        pdf.set_font("Arial", 'B', 14)  # Negrita, tamaño 14
        pdf.cell(0, 10, f"Sesión {i}", ln=True, align='L')
        pdf.set_font("Arial", size=12)  # Normal para los grupos
        for j, grupo in enumerate(sesion, 1):
            nombres_completos = [f"{alumnos[a][0]} {alumnos[a][1]} {alumnos[a][2]}" for a in grupo]
            texto_grupo = f"Grupo {j}: " + ", ".join(nombres_completos)
            pdf.cell(0, 8, texto_grupo, ln=True)
        pdf.ln(5)
    pdf.output(nombre_archivo)

# === Programa principal ===
if __name__ == "__main__":
    df = pd.read_excel(RUTA_EXCEL)
    alumnos = list(zip(df[COL_NOMBRE], df[COL_APELLIDO1], df[COL_APELLIDO2]))
    sesiones = generar_sesiones(alumnos, SESIONES_MAXIMAS, INTEGRANTES_POR_GRUPO, INTENTOS_POR_SESION)
    print(f"Sesiones generadas: {len(sesiones)}")
    exportar_pdf(sesiones, alumnos, INTEGRANTES_POR_GRUPO)
    print("PDF generado")