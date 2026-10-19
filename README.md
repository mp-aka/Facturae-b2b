# Facturae B2B · Conversor Oficial Facturae 3.2.2

Aplicación web profesional para la generación y validación de facturas electrónicas en formato oficial **Facturae 3.2.2**, conforme a la **Ley Crea y Crece** y los requisitos de presentación ante **FACe** y relaciones B2B en España.

---

## 🚀 Características Principales

- **Procesamiento 100% en Memoria Volátil (RAM):** Ningún dato fiscal de emisores, receptores o líneas de factura se almacena en el disco del servidor ni en bases de datos externas.
- **Validación Fiscal Estricta:** Comprobación sintáctica y aritmética de NIF/CIF, tipos impositivos (IVA/IRPF), IBAN bancario y códigos DIR3 de la Administración Pública.
- **Integración con AutoFirma:** Detección automática en escritorio local (macOS, Windows, Linux) para firma con certificados FNMT/DNIe o descarga directa del XML.
- **Pasarela de Pagos Stripe Nativa:**
  - **Suscripción Mensual:** 9,49 € / mes (Facturación flexible con portal de autoservicio para clientes).
  - **Licencia PRO Vitalicia:** 99,49 € (Pago único con acceso permanente e ilimitado de por vida).
- **Procesador de Webhooks Seguro:** Manejo idempotente de eventos de facturación (`checkout.session.completed`, `customer.subscription.deleted`, `invoice.payment_succeeded`, `invoice.payment_failed`).
- **Blindaje Red Team:** Cero exposición de API keys, contraseñas cifradas con PBKDF2-SHA256 (100.000 iteraciones), mitigación de CSRF/XSRF y bloqueo de descargas para usuarios sin licencia activa.

---

## 📦 Requisitos del Sistema

- **Python:** 3.10 o superior (compatible con 3.11, 3.12, 3.14).
- **Dependencias principales:** `streamlit`, `pandas`, `openpyxl`, `stripe`.

---

## ⚡ Instalación y Ejecución Local

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/tu-usuario/FacturaeB2B.git
   cd FacturaeB2B
   ```

2. **Crear y activar un entorno virtual:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # En Windows: .venv\Scripts\activate
   ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configurar secretos de prueba:**
   ```bash
   cp .streamlit/secrets.toml.example .streamlit/secrets.toml
   ```
   *(Edita `.streamlit/secrets.toml` con tus claves de Stripe de prueba y ajusta `MODO_DESARROLLO = "true"`).*

5. **Lanzar la aplicación:**
   ```bash
   streamlit run app.py
   ```
   Abre tu navegador en `http://localhost:8501`.

---

## 🌐 Despliegue en Producción (Streamlit Community Cloud)

1. Sube tu código a un repositorio privado en **GitHub**.
2. Entra en **[share.streamlit.io](https://share.streamlit.io/)** e inicia sesión con GitHub.
3. Haz clic en **«New app»** y selecciona:
   - **Repository:** `tu-usuario/FacturaeB2B`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. En la configuración de la app (**Settings > Secrets**), copia el contenido de `.streamlit/secrets.toml.example` rellenando tus datos de producción:
   - `MODO_DESARROLLO = "false"`
   - `APP_BASE_URL = "https://tu-app.streamlit.app"`
   - Claves de Stripe (`STRIPE_SECRET_KEY`, `STRIPE_PUBLISHABLE_KEY`)
   - Credenciales SMTP para los correos de verificación.
5. Haz clic en **Deploy**. Tu pasarela estará operativa con HTTPS oficial en menos de 2 minutos.

---

## 🔒 Auditoría de Seguridad

- **API Keys y Credenciales:** Protegidas bajo `.gitignore`. Ninguna clave privada o contraseña se versiona en Git.
- **Inyección SQL:** Todas las operaciones de base de datos utilizan parámetros vinculados (`?`).
- **Inyección de Fórmulas:** Los registros de soporte neutralizan caracteres de control en CSV (`=`, `+`, `-`, `@`).
- **Validación Criptográfica de Webhooks:** Firma HMAC-SHA256 con ventana de tolerancia temporal contra ataques de repetición.

---

## ⚖️ Licencia y Aviso Legal

Esta herramienta realiza la conversión sintáctica de datos en memoria para generar el formato Facturae 3.2.2. No constituye asesoría contable ni fiscal directa.
