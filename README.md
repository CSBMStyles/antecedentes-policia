# Antecedentes Judiciales - Policía Nacional

Automatiza la consulta de antecedentes judiciales en el sitio oficial de la
Policía Nacional de Colombia:

`https://antecedentes.policia.gov.co:7005/WebJudicial/`

El **CAPTCHA se resuelve manualmente**: la aplicación abre Chrome, llena la
cédula y espera. Si Google muestra la cuadrícula de imágenes, tú seleccionas las
imágenes y pulsas **Verificar**; la aplicación detecta el token y pulsa
**Consultar** sola.

---

## Qué hace

- Lee los empleados desde `CARACTERIZACION.xlsx` (hoja `CONVENIO EJECUCION`).
- Respeta **exactamente el orden de las filas del Excel** (la base está ordenada
  por apellido).
- Permite elegir hoja, fila de encabezados y qué columna tiene cada dato:
  documento, nombres, apellidos, perfil y tipo de documento.
- Abre Chrome una sola vez para todo el lote.
- Consulta cada cédula y guarda el certificado en PDF:
  `ANTECEDENTES POLICIA/{PERFIL}/APELLIDO NOMBRE - CEDULA.pdf`
  (solo la primera página: el certificado; la segunda página del sitio, que
  únicamente tiene el pie de página, se descarta).
- Genera `REPORTE_FALLOS_<fecha>.xlsx`, `.csv` y `.txt` con un resumen claro.
- Guarda el avance en `estado_antecedentes_policia.json` para reanudar.
- Permite trabajar por tandas y por perfil.

---

## El CAPTCHA manual, paso a paso

1. La aplicación abre Chrome y acepta los términos del sitio.
2. Escribe la cédula en el formulario.
3. Intenta marcar el checkbox de reCAPTCHA.
4. Si Google muestra la cuadrícula (“selecciona todas las imágenes de
   autobús”, por ejemplo), **tú seleccionas las imágenes y pulsas Verificar**.
5. Cuando el token queda listo, la aplicación pulsa **Consultar**.
6. Espera el certificado, valida que sea el documento correcto y lo guarda.

No necesitas API keys ni pagar 2captcha. Si el checkbox pasa solo, no tienes que
hacer nada.

> El tiempo de espera se configura en la pestaña `3 · CAPTCHA manual`
> (300 segundos por defecto). Si se agota, esa persona queda en el reporte y la
> puedes reintentar después.

---

## Instalación

### macOS

```bash
cd Codigo/ANTECEDENTES/obtener-antecedentes-policia-main
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Abre la interfaz con doble clic en **`iniciar.command`** o:

```bash
python app.py
```

Si macOS bloquea el chromedriver:

```bash
xattr -d com.apple.quarantine "$HOME/Library/Application Support/undetected_chromedriver/undetected_chromedriver"
```

### Windows

1. Instala **Python 3.11 o superior** y marca **“Add python.exe to PATH”**.
2. Instala **Google Chrome**.
3. En PowerShell:

```powershell
cd Codigo\ANTECEDENTES\obtener-antecedentes-policia-main
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Si PowerShell bloquea la activación:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

4. Abre la interfaz con doble clic en **`iniciar.bat`** o:

```powershell
python app.py
```

---

## La interfaz

### 1 · Datos

- Excel de caracterización y hoja.
- Fila de encabezados.
- Mapeo de columnas: documento, nombres, apellidos, perfil y tipo de documento
  (este último es opcional; si no aparece, se usa cédula de ciudadanía).
- Vista previa con las columnas mapeadas primero.
- **Validar datos**: cuenta válidos, incompletos y duplicados sin abrir Chrome.

### 2 · Salida y navegador

- Carpeta base (`ANTECEDENTES POLICIA` por defecto).
- Nombre del PDF:
  - `APELLIDO NOMBRE - CEDULA` (recomendado, conserva el orden de la base).
  - `CEDULA NOMBRE APELLIDO`.
- Chrome siempre visible: es obligatorio para resolver el captcha.
- Driver: `auto`, `uc` o `selenium`.

### 3 · CAPTCHA manual

- Instrucciones y tiempo de espera.
- **Probar CAPTCHA**: abre Chrome, muestra el formulario y espera a que resuelvas
  un captcha. No consulta ninguna cédula; solo confirma que el flujo funciona.

### 4 · Ejecución

- Perfil (`TODOS`, `AUXILIAR`, `ENFERMERO`, `MEDICO`).
- Cantidad por tanda (`0` = todos).
- Reanudar o reiniciar el estado.
- Intentos por persona si el sitio rechaza el captcha.
- Barra de progreso y botón **Detener**.

---

## Salida

```
ANTECEDENTES POLICIA/
├── AUXILIAR/
│   ├── PEREZ GOMEZ ANA MARIA - 1000000001.pdf
│   └── ...
├── ENFERMERO/
└── MEDICO/
```

El nombre empieza por el apellido para que la carpeta quede en el mismo orden
alfabético de la base de datos.

---

## Uso por línea de comandos

```bash
python main.py
python main.py --perfil AUXILIAR --tanda 10 --si
python main.py --perfil TODOS --si
python main.py --reiniciar-estado --si
python main.py --excel "/ruta/CARACTERIZACION.xlsx" --salida "/ruta/salida"
python main.py --hoja "CONVENIO EJECUCION" --fila-encabezado 1 --si
python main.py --dry-run --si
python main.py --timeout-manual 600 --si
```

`--dry-run` valida el Excel sin abrir Chrome.

---

## Problemas frecuentes

### Chrome no aparece

Instala Google Chrome (no basta Edge). La pestaña 2 muestra `Chrome: OK` con la
ruta detectada.

### “This version of ChromeDriver only supports Chrome version …”

`undetected-chromedriver` tiene un driver viejo. El programa cae automáticamente a
**Selenium estándar**, que descarga el driver correcto. Si quieres forzarlo:

```bash
DRIVER_BACKEND=selenium
```

o elige `selenium` en la pestaña 2.

### El checkbox pasa solo

Perfecto: no tendrás que resolver la cuadrícula. La aplicación lo detecta y sigue.

### Aparece la cuadrícula pero no avanza

- Asegúrate de pulsar **Verificar** después de seleccionar las imágenes.
- Si el sitio pide otra ronda, complétala.
- Si el token no aparece, revisa que Chrome tenga el foco y que no haya otra
  pestaña bloqueando.
- Aumenta el tiempo de espera en la pestaña 3.

### El PDF no se guardó

- Revisa el reporte de fallos y la captura en `errors/`.
- Verifica que la carpeta de salida tenga permisos de escritura.
- Si el sitio devolvió la página de términos, el captcha fue rechazado: la
  aplicación lo marca y reintenta según los intentos configurados.

### Windows

- Si el antivirus bloquea `undetected_chromedriver.exe`, usa `selenium` como
  driver.
- Si el proyecto está en OneDrive, mueve la carpeta de salida a una ruta local.
- Los nombres se recortan a 120 caracteres para no superar el límite de Windows.

---

## Archivos

- `app.py`: lanzador de la interfaz.
- `gui.py`: ventana.
- `main.py`: núcleo de consulta y CLI.
- `config.py`: `config.json` y `.env`.
- `sources.py`: lectura del Excel, hojas, mapeo y orden.
- `utils.py`: texto, tipos de documento y nombres seguros.
- `perfiles.py`: campos del proyecto (base para los otros antecedentes).
- `estado_antecedentes_policia.json`: control de avance.
- `errors/`: capturas cuando falla una consulta.
- `REPORTE_FALLOS_*.xlsx`: reporte en tabla, con una fila por persona.
- `REPORTE_FALLOS_*.csv` / `.txt`: los mismos datos en texto.

---

## Reutilizar en los otros antecedentes

`perfiles.py` declara los campos y el tipo de captcha. Para procuraduría, rethus,
inhabilidades, etc., se copia la carpeta y se cambia:

- `PROYECTO` (nombre y entidad).
- `CAMPOS` (qué columnas necesita ese trámite).
- El flujo del sitio en `main.py`.

La ventana, `config.py`, `sources.py` y el manejo de estado/reportes se reutilizan.

---

## Reporte de fallos

El archivo `.xlsx` tiene dos hojas:

- **Fallos**: una fila por persona con estas columnas:
  `Fila Excel`, `Documento`, `Nombre`, `Perfil`, `Motivo`, `Qué pasó`,
  `Qué hacer`, `Mensaje del sitio`, `Intentos`, `Método`, `Captura`.
- **Resumen**: cuántos fallos hay por motivo y qué significa cada uno.

El `.txt` es la versión corta para leer rápido y el `.csv` sirve para filtrar en
Excel. La captura de cada fallo queda enlazada desde la columna `Captura`.

---

## Uso autorizado

Este proyecto está pensado para que una entidad autorizada consulte los
antecedentes de sus colaboradores. El sitio de la Policía Nacional restringe el
uso de la información a fines personales o institucionales legítimos; no la
compartas ni la uses para otros fines.

---

## Datos y privacidad

Este repositorio contiene únicamente código y documentación. **No incluye**:

- API keys ni archivos `.env`.
- Bases de datos de empleados, cédulas, nombres, correos o teléfonos.
- PDFs, Excel, PPTX, CSV, capturas o reportes generados por las consultas.
- Configuraciones locales (`config.json`) ni estados de proceso.

Para usarlo, proporciona tu propia base de datos autorizada y configura tus
propias credenciales. No subas datos personales al repositorio.

El archivo `config.example.json` muestra una configuración mínima; cópialo a
`config.json` o deja que la interfaz lo cree.

## Licencia

MIT. Consulta el archivo `LICENSE`.

