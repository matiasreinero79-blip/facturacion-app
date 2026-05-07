# Sistema de Gestión de Facturas

App de escritorio en Python/tkinter para automatizar la descarga, organización y control de facturas.
**Compatible con Windows, macOS y Linux.**

---

## Estructura del proyecto

```
facturacion_app/
├── main.py                        ← Punto de entrada
├── app.py                         ← GUI principal (tkinter)
├── config.py                      ← Carga/guarda config.json
├── config.json                    ← Creado automáticamente al guardar ajustes
├── control_facturas.xlsx          ← Creado automáticamente al procesar facturas
├── Facturas/                      ← Creado automáticamente
│   └── 2026-05/                   ← Sub-carpetas por año-mes
├── requirements.txt
└── modules/
    ├── utils.py                   ← Helpers cross-platform (open_path, fonts, themes)
    ├── pdf_reader.py              ← Extrae texto de PDFs
    ├── file_manager.py            ← Copia y organiza archivos
    ├── excel_manager.py           ← Lee/escribe control_facturas.xlsx
    ├── email_sender.py            ← Envía email vía SMTP
    └── importers/
        ├── base_importer.py       ← Contrato base
        ├── folder_importer.py     ← Importar desde carpeta local ✔
        ├── email_importer.py      ← Placeholder (próximamente)
        └── web_importer.py        ← Placeholder (próximamente)
```

---

## Requisitos

- **Python 3.9 o superior**
- tkinter incluido con Python (ver nota por OS más abajo)

---

## Instalación

### Windows

```powershell
cd facturacion_app
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Si PowerShell bloquea la activación del venv, ejecuta primero (una sola vez):
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

### macOS

```bash
cd facturacion_app
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

> Si `tkinter` no está disponible (`ModuleNotFoundError: No module named 'tkinter'`),
> instala Python desde [python.org](https://www.python.org/downloads/) (el de la Mac
> App Store y el del sistema no siempre incluyen tkinter).
> Alternativamente con Homebrew: `brew install python-tk`

### Linux (Debian / Ubuntu / Mint)

```bash
# Instalar tkinter si no está incluido
sudo apt install python3-tk

cd facturacion_app
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

> En Fedora/RHEL: `sudo dnf install python3-tkinter`
> En Arch: `sudo pacman -S tk`

---

## Ejecutar la app

```bash
# Windows (PowerShell)
python main.py

# macOS / Linux
python3 main.py
```

---

## Uso rápido

1. **Seleccionar carpeta** → elige la carpeta donde están tus facturas PDF.
2. **Buscar PDFs** → la lista se llena con los archivos encontrados.
3. **Procesar** → por cada PDF se abre un formulario pre-rellenado con los datos extraídos automáticamente. Revisa, completa y guarda.
   - El PDF se copia a `Facturas/YYYY-MM/`.
   - El registro se agrega a `control_facturas.xlsx`.
   - Si la factura ya existe (misma empresa + número de factura), se avisa y se omite.
4. **Abrir Excel** → abre `control_facturas.xlsx` con Excel o LibreOffice Calc.
5. **Enviar Email** → envía un resumen HTML con las facturas Pendientes al área de finanzas.
   Configura primero el email en **Configuración Email**.

### Cambiar estado de una factura

En la pestaña **Control de Facturas**:
- **Windows / Linux**: clic derecho sobre una fila.
- **macOS**: `Ctrl + clic` sobre una fila.

Opciones: *Marcar como Paga* / *Marcar como Pendiente* / *Abrir PDF*.

---

## Configuración de email

Abre **Configuración Email** en la cabecera de la app.

| Campo | Descripción |
|-------|-------------|
| Email Remitente | Tu dirección de correo |
| Email Finanzas | Destinatario del aviso |
| Contraseña / App Password | Para Gmail, genera una *App Password* en la seguridad de tu cuenta Google (no uses tu contraseña normal) |
| Servidor SMTP | `smtp.gmail.com` (Gmail), `smtp.office365.com` (Outlook/Microsoft 365) |
| Puerto SMTP | `587` (STARTTLS) o `465` (SSL) |
| Usar TLS | Activo por defecto |

La configuración se guarda en `config.json`. **No compartas este archivo** ya que contiene la contraseña en texto plano.

---

## Generar ejecutable

### Windows — `.exe` con PyInstaller

```powershell
pip install pyinstaller
pyinstaller --onefile --windowed --name FacturacionApp main.py
```

El ejecutable queda en `dist\FacturacionApp.exe`.

### macOS — `.app` con PyInstaller

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name FacturacionApp main.py
```

El bundle queda en `dist/FacturacionApp.app`.

### Linux — binario con PyInstaller

```bash
pip install pyinstaller
pyinstaller --onefile --name FacturacionApp main.py
```

El binario queda en `dist/FacturacionApp`.

> **Nota antivirus**: los binarios generados con PyInstaller a veces son
> detectados como falsos positivos. Agrégalos a la lista de exclusiones si
> es necesario.

---

## Fuentes de importación futuras

| Módulo | Estado |
|--------|--------|
| `FolderImporter` — carpeta local | ✔ Operativo |
| `EmailImporter` — IMAP | Próximamente |
| `WebPortalImporter` — scraping/API | Próximamente |

Para agregar una nueva fuente, crear una clase que herede `BaseImporter`
e implemente `get_invoices() -> List[Path]`.
