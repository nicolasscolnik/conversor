# 🧾 Conversor de Facturas a Excel

**Conversor de Facturas** es una aplicación de escritorio que permite extraer automáticamente información de facturas en PDF y generar un reporte en Excel (.xlsx) usando **Inteligencia Artificial (OpenAI GPT-4o)**.

## ✅ Características principales

- Extrae de cada factura:
    - Fecha
    - Tipo de Factura (A, B, C, M, E, T)
    - Número de Factura
    - Nombre del Vendedor
    - Monto Total
- Guarda los resultados en un Excel organizado, con detección automática de errores.

---

## 🚀 Funcionalidades principales

- 📄 Selección masiva de facturas en PDF.
- 🔥 Procesamiento automático usando **OpenAI GPT-4o**.
- 📊 Generación de un Excel listo para su uso.
- 🚦 Barra de progreso en tiempo real.
- 🛡️ Seguridad: la API Key se maneja desde archivo `.env`.
- 🗂️ No sobrescribe archivos: si el Excel ya existe, crea `(+1)` automáticamente.
- ✨ Interfaz gráfica moderna usando **CustomTkinter**.

---

## 🛠️ Requisitos

- Python 3.8 o superior.
- Cuenta activa en [OpenAI](https://platform.openai.com/).
- API Key de OpenAI.
- Google Drive configurado para autenticación inicial.

### Instalación de dependencias

```bash
pip install -r requirements.txt
```

### ⚙️ Configuración inicial

Crear un archivo `.env` en el proyecto raíz con el siguiente contenido:

```env
OPENAI_API_KEY=tu-api-key-de-openai
```

Tener disponible el archivo `credentials.json` para Google Drive OAuth (lo puedes obtener desde Google Cloud Console).

---

## 📂 Estructura del proyecto

```plaintext
CONVERSOR/
├── Facturas/                  # Carpeta opcional para PDFs
├── build/                     # Carpeta de build si compilas el .exe
├── dist/                      # Carpeta de distribución
├── conversor_facturas_v3.0.py # Código fuente principal
├── .env                       # Variables de entorno (NO subir a GitHub)
├── client_secrets.json        # Secretos de Drive (NO subir a GitHub)
├── credentials.json           # Credenciales de Drive generadas
├── requirements.txt
└── README.md
```

---

## 🖥️ Cómo ejecutarlo en modo local

```bash
python conversor_facturas_v3.0.py
```

Se abrirá una ventana de escritorio donde podrás:

1. Seleccionar una carpeta de facturas.
2. Procesarlas automáticamente.
3. Descargar el archivo Excel generado.

---

## 📈 Ejemplo de Excel generado

| Fecha       | Tipo de Factura | Número de Factura | Nombre del Vendedor         | Monto Total |
|-------------|-----------------|-------------------|-----------------------------|-------------|
| 24/04/2025  | B               | 0007-00004076     | Andrés Rodríguez Grumberg   | 22658.19    |
| 15/04/2025  | B               | 0011-00664426     | Electropoint                | 68549.00    |
| 27/03/2025  | B               | 00003-00050655    | Rutina Mercado Libre        | 29729.00    |
| 22/04/2025  | B               | 00021-02265704    | Brandlive S.A.              | 87999.20    |

---

## ⚡ Compilar a .exe (opcional)

Compilar la aplicación para Windows:

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole conversor_facturas_v3.0.py
```

El ejecutable aparecerá en la carpeta `/dist`.

> **Nota:** El archivo `.env` y `credentials.json` deben estar en la misma carpeta que el `.exe` si quieres distribuirlo.

---

## 🛡️ Consideraciones de Seguridad

- Tu `OPENAI_API_KEY` nunca debe subirse al repositorio público.
- Usa siempre un archivo `.gitignore` para ignorar `.env` y `credentials.json`.
- No compartas públicamente las claves o secretos.

---

## 👨‍💻 Autor

Proyecto desarrollado por:

**Nicolás Scolnik**

📬 Contacto:  
- [LinkedIn](https://www.linkedin.com/in/nicolas-scolnik-it/)  
- [Email](nicolasscolnik@gmail.com)

---

## 📝 Licencia

Este proyecto es privado y de uso exclusivo para fines educativos o internos.  
Está prohibido el uso comercial sin autorización previa.

---

🎯 ¡Ahora puedes automatizar la generación de reportes de facturas con Inteligencia Artificial!  