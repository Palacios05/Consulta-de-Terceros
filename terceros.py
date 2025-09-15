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
ruta_excel = r"/home/palaciosjj/Documents/Trabajo TDM/Consulta-de-Terceros/terceros.xlsx"
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
def resolver_pregunta(pregunta: str, cedula: str) -> str:
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

    # --- Preguntas de dígitos de la cédula ---
    if "primeros" in texto and "dig" in texto:
        nums = re.findall(r"\d+", texto)
        if nums:
            n = int(nums[0])
            return cedula[:n]
    if "ultim" in texto and "dig" in texto:
        nums = re.findall(r"\d+", texto)
        if nums:
            n = int(nums[0])
            return cedula[-n:]

    # --- Preguntas de cultura general ---
    for clave, respuesta in preguntas_conocidas.items():
        if clave in texto:
            return respuesta

    # --- Fallback ---
    return "no se"

def procesar_nombre(nombre_completo: str):
    partes = [p for p in nombre_completo.split() if p.strip()]

    if len(partes) == 0:
        return "", "", "", ""
    if len(partes) == 1:
        return partes[0], "", "", ""
    if len(partes) == 2:
        return partes[0], "", partes[1], ""
    if len(partes) == 3:
        return partes[0], "", partes[1], partes[2]
    if len(partes) == 4:
        return partes[0], partes[1], partes[2], partes[3]
    primer_nombre = partes[0]
    segundo_nombre = " ".join(partes[1:-2]) if len(partes) > 3 else ""
    return primer_nombre, segundo_nombre, partes[-2], partes[-1]

def extraer_nombre_desde_div(datos_div):
    try:
        spans = datos_div.find_elements(By.TAG_NAME, "span")
        names = []
        for s in spans:
            txt = (s.get_attribute("textContent") or "").strip()
            if txt:
                names.append(txt)
        if names:
            return " ".join(names).strip()

        for tag in ("strong", "b", "em", "i"):
            elems = datos_div.find_elements(By.TAG_NAME, tag)
            texts = [e.text.strip() for e in elems if e.text and e.text.strip()]
            if texts:
                return " ".join(texts).strip()

        inner = datos_div.get_attribute("innerText") or datos_div.text or ""
        texto = re.sub(r"\s+", " ", inner).strip()
        if not texto:
            return ""

        lower = texto.lower()

        if "señor" in lower or "senor" in lower:
            idx = lower.find("señor") if "señor" in lower else lower.find("senor")
            sub = texto[idx:]
            sub = re.sub(r"(?i)señor\(a\)|(?i)señor", "", sub, count=1).strip()
            markers = ["identificado", "identificada", "con cédula", "con cedula", "cédula", "cedula", "número", "numero"]
            end_pos = len(sub)
            for m in markers:
                p = sub.lower().find(m)
                if p != -1 and p < end_pos:
                    end_pos = p
            candidate = sub[:end_pos].strip(" :.,")
            candidate = re.sub(r"[^\w\sÁÉÍÓÚÑÜáéíóúñüÀ-ÿ]", "", candidate).strip()
            if candidate:
                return candidate

        upper_matches = re.findall(r"([A-ZÁÉÍÓÚÑÜ]{2,}(?:\s+[A-ZÁÉÍÓÚÑÜ]{2,})+)", texto)
        if upper_matches:
            return max(upper_matches, key=len)

        lines = [ln.strip() for ln in inner.splitlines() if ln.strip()]
        for ln in lines:
            if re.search(r"\d", ln):
                continue
            if re.search(r"cedula|identificad|c.i.|numero", ln, flags=re.I):
                continue
            if len(ln.split()) >= 2:
                return re.sub(r"[^\w\sÁÉÍÓÚÑÜáéíóúñüÀ-ÿ]", "", ln).strip()

        fallback = re.sub(r"[^\w\sÁÉÍÓÚÑÜáéíóúñüÀ-ÿ]", " ", texto).strip()
        return fallback

    except Exception:
        return ""

# ====================================================
# MAIN
# ====================================================
wb = openpyxl.load_workbook(ruta_excel)
hoja = wb.active

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

for fila in range(2, hoja.max_row + 1):
    cedula = hoja[f"A{fila}"].value
    if cedula is None:
        continue
    cedula = str(cedula).strip()

    try:
        driver.get(url)
        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)

        tipo_id = wait.until(EC.presence_of_element_located((By.ID, "ddlTipoID")))
        Select(tipo_id).select_by_value("1")

        campo_cedula = wait.until(EC.presence_of_element_located((By.NAME, "txtNumID")))
        campo_cedula.clear()
        campo_cedula.send_keys(cedula)

        exito = False
        for intento in range(3):
            pregunta_elem = wait.until(EC.presence_of_element_located((By.ID, "lblPregunta")))
            pregunta = pregunta_elem.text.strip()
            respuesta = resolver_pregunta(pregunta, cedula)

            if respuesta == "no se":
                refresh_btn = wait.until(EC.element_to_be_clickable((By.ID, "ImageButton1")))
                refresh_btn.click()
                time.sleep(2)
                continue

            campo_respuesta = wait.until(EC.presence_of_element_located((By.NAME, "txtRespuestaPregunta")))
            campo_respuesta.clear()
            campo_respuesta.send_keys(str(respuesta))

            btn_consultar = wait.until(EC.element_to_be_clickable((By.ID, "btnConsultar")))
            btn_consultar.click()

            try:
                try:
                    datos_div = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.ID, "datosConsultado"))
                    )
                except Exception:
                    datos_div = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "datosConsultado"))
                    )

                nombre_completo = extraer_nombre_desde_div(datos_div).strip()

                if not nombre_completo:
                    hoja[f"G{fila}"] = "Fallo"
                    wb.save(ruta_excel)
                    break

                primer_nombre, segundo_nombre, primer_apellido, segundo_apellido = procesar_nombre(nombre_completo)

                hoja[f"B{fila}"] = primer_nombre
                hoja[f"C{fila}"] = segundo_nombre
                hoja[f"D{fila}"] = primer_apellido
                hoja[f"E{fila}"] = segundo_apellido
                hoja[f"F{fila}"] = nombre_completo
                hoja[f"G{fila}"] = "Éxito"

                print(f"[OK] Cédula {cedula} -> {nombre_completo}")
                exito = True
                wb.save(ruta_excel)
                break

            except Exception as e:
                print(f"[SIN DATOS] No se pudo obtener información para la cédula {cedula}: {e}")
                hoja[f"G{fila}"] = "Fallo"
                wb.save(ruta_excel)
                break

        if not exito:
            hoja[f"G{fila}"] = "Fallo"
            wb.save(ruta_excel)

    except Exception as e:
        print(f"[ERROR] con la cédula {cedula}: {e}")
        hoja[f"G{fila}"] = "Fallo"
        wb.save(ruta_excel)

    finally:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

driver.quit()
