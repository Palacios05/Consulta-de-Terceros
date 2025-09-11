import time
import re
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ====================================================
# CONFIGURACIÓN
# ====================================================
ruta_excel = r"C:\Users\Usuario\Documents\Scripts consultar terceros\terceros.xlsx"
url = "https://www.procuraduria.gov.co/Pages/Consulta-de-Antecedentes.aspx"

# Diccionario de preguntas de cultura general
preguntas_conocidas = {
    "capital del valle del cauca": "cali",
    "capital de colombia": "bogota",
    "capital de antioquia": "medellin",
    "capital del atlantico": "barranquilla",
    "capital del huila": "neiva",
    "capital del tolima": "ibague",
    "capital del cauca": "popayan",
    "capital del nariño": "pasto",
    "capital del meta": "villavicencio",
    "oceanos que limitan con colombia": "pacifico y atlantico",
    "continente donde esta colombia": "america",
    "color de la bandera de colombia": "amarillo azul y rojo",
    "capital del quindio": "armenia",
    "capital de risaralda": "pereira"
}

# ====================================================
# FUNCIONES
# ====================================================
def resolver_pregunta(pregunta: str) -> str:
    texto = pregunta.lower().strip()
    numeros = re.findall(r"\d+", texto)

    # --- Resolver matemáticas ---
    if numeros:
        numeros = list(map(int, numeros))
        if "mas" in texto or "+" in texto or "sume" in texto or "suma" in texto:
            return str(sum(numeros))
        elif "menos" in texto or "-" in texto or "reste" in texto:
            return str(numeros[0] - numeros[1])
        elif "x" in texto or "multiplique" in texto or "por" in texto:
            return str(numeros[0] * numeros[1])
        elif "divida" in texto or "/" in texto:
            return str(numeros[0] // numeros[1])

    # --- Preguntas de cultura general ---
    for clave, respuesta in preguntas_conocidas.items():
        if clave in texto:
            return respuesta

    # --- Fallback ---
    return "no se"

def procesar_nombre(nombre_completo: str):
    partes = nombre_completo.split()

    if len(partes) == 4:  # Ej: Juan Carlos Pérez Gómez
        return partes[0], partes[1], partes[2], partes[3]
    elif len(partes) == 3:  # Ej: Juan Pérez Gómez
        return partes[0], "", partes[1], partes[2]
    elif len(partes) == 2:  # Ej: Juan Pérez
        return partes[0], "", partes[1], ""
    elif len(partes) > 4:  # Ej: María del Carmen López Pérez
        primer_nombre = partes[0]
        segundo_nombre = " ".join(partes[1:-2])
        return primer_nombre, segundo_nombre, partes[-2], partes[-1]
    else:  # Solo un nombre
        return partes[0], "", "", ""

# ====================================================
# MAIN
# ====================================================
# Abrir Excel
wb = openpyxl.load_workbook(ruta_excel)
hoja = wb.active

# Configurar Selenium
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)  # espera hasta 20 segundos para elementos

for fila in range(2, hoja.max_row + 1):  # Empieza en fila 2
    cedula = hoja[f"A{fila}"].value  # Columna A = NIT Receptor
    if cedula is None:
        continue

    try:
        driver.get(url)

        # --- Cambiar al iframe ---
        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)

        # Seleccionar tipo de documento = Cédula
        tipo_id = wait.until(EC.presence_of_element_located((By.ID, "ddlTipoID")))
        Select(tipo_id).select_by_value("1")

        # Escribir cédula
        campo_cedula = wait.until(EC.presence_of_element_located((By.NAME, "txtNumID")))
        campo_cedula.clear()
        campo_cedula.send_keys(str(cedula))

        # Leer y resolver pregunta
        pregunta_elem = wait.until(EC.presence_of_element_located((By.ID, "lblPregunta")))
        pregunta = pregunta_elem.text.strip()
        respuesta = resolver_pregunta(pregunta)

        # Escribir respuesta
        campo_respuesta = wait.until(EC.presence_of_element_located((By.NAME, "txtRespuestaPregunta")))
        campo_respuesta.clear()
        campo_respuesta.send_keys(str(respuesta))

        # Enviar formulario
        btn_consultar = wait.until(EC.element_to_be_clickable((By.ID, "btnConsultar")))
        btn_consultar.click()

        # --- Esperar hasta 30 segundos la respuesta ---
        try:
            datos_div = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "datosConsultado"))
            )
            spans = datos_div.find_elements(By.TAG_NAME, "span")
            nombre_completo = " ".join([s.text.strip() for s in spans if s.text.strip()])

            primer_nombre, segundo_nombre, primer_apellido, segundo_apellido = procesar_nombre(nombre_completo)

            # Guardar en Excel (ajusta columnas según tu formato)
            hoja[f"B{fila}"] = primer_nombre
            hoja[f"C{fila}"] = segundo_nombre
            hoja[f"D{fila}"] = primer_apellido
            hoja[f"E{fila}"] = segundo_apellido

            print(f"[OK] Cédula {cedula} -> {nombre_completo}")

        except Exception:
            print(f"[SIN DATOS] No se pudo obtener información para la cédula {cedula}")

    except Exception as e:
        print(f"[ERROR] con la cédula {cedula}: {e}")

    finally:
        # volver al contexto principal antes de la siguiente iteración
        driver.switch_to.default_content()

# Guardar Excel
wb.save(ruta_excel)
driver.quit()
