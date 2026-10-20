import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from io import BytesIO
from datetime import date, datetime, timedelta
from pathlib import Path
import subprocess
import platform
import shutil
import os
import re
import csv
import sqlite3
import hashlib
import secrets
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

try:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    import subprocess
    import sys
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl>=3.1.2", "stripe>=11.0.0"])
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.worksheet.datavalidation import DataValidation
    except Exception:
        pass


try:
    import stripe
    import managed_payments
    import stripe_billing
except ImportError:
    stripe = None
    managed_payments = None
    stripe_billing = None

# ==============================================================================
# CONFIGURACIÓN GENERAL Y HARDENING DE SEGURIDAD
# ==============================================================================

st.set_page_config(
    page_title="Facturae B2B · Conversor Seguro Excel a XML 3.2.2",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Límite estricto de subida en RAM (5 MB) para mitigar denegación de servicio (DoS)
MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024

# Rutas locales
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "facturae_b2b.db"
LOG_ERRORES_PATH = BASE_DIR / "logs_errores_usuarios.csv"
COLUMNAS_LOG_ERRORES = ["Fecha_Hora", "Categoria", "Descripcion_Usuario"]
CATEGORIAS_REPORTE = [
    "Selecciona el tipo de error...",
    "El Excel da error al subir",
    "Cálculo de impuestos incorrecto",
    "Problema con la plantilla",
    "AutoFirma no se abre",
    "Problema con verificación de correo",
    "Problema con mi cuenta o código",
    "Otro problema",
]

# Lista negra de proveedores de correos desechables o temporales (anti-spam / antifraude)
DOMINIOS_DESECHABLES = {
    "tempmail.com", "guerrillamail.com", "yopmail.com", "mailinator.com",
    "10minutemail.com", "trashmail.com", "sharklasers.com", "getairmail.com",
    "dispostable.com", "temp-mail.org", "fakeinbox.com", "mytemp.email",
    "mohmal.com", "crazymailing.com", "inboxkitten.com", "throwawaymail.com"
}


# ==============================================================================
# GESTIÓN SEGURA DE CLAVES (CERO HARDCODING - st.secrets)
# ==============================================================================

def obtener_secreto(clave: str, valor_defecto: str = "") -> str:
    """Obtiene valores confidenciales exclusivamente desde st.secrets."""
    try:
        if clave in st.secrets:
            return str(st.secrets[clave]).strip()
    except Exception:
        pass
    return valor_defecto


# Entorno de ejecución: en producción MODO_DESARROLLO debe ser False
MODO_DESARROLLO = obtener_secreto("MODO_DESARROLLO", "false").lower() in ("true", "1", "yes")

STRIPE_SECRET_KEY = obtener_secreto("STRIPE_SECRET_KEY")
STRIPE_PUBLISHABLE_KEY = obtener_secreto("STRIPE_PUBLISHABLE_KEY")
STRIPE_PRICE_VITALICIA = obtener_secreto("STRIPE_PRICE_VITALICIA", "price_1USGYQROobv6Wd8mcdXDWFl1")
STRIPE_PRICE_MENSUAL = obtener_secreto("STRIPE_PRICE_MENSUAL", "price_1USGYQROobv6Wd8mr2jwehcN")
APP_BASE_URL = obtener_secreto("APP_BASE_URL", "http://localhost:8501")
STRIPE_PAYMENT_URL = obtener_secreto(
    "STRIPE_PAYMENT_URL",
    "https://buy.stripe.com/test_facturaeb2b_licencia"
)


# ==============================================================================
# BASE DE DATOS LOCAL SEGURA (USUARIOS, VERIFICACIÓN Y CÓDIGOS DE UN SOLO USO)
# ==============================================================================

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Inicializa las tablas de usuarios, verificaciones OTP, transacciones Stripe y códigos de un solo uso en SQLite con modo WAL."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("PRAGMA synchronous=NORMAL;")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                email_verificado INTEGER DEFAULT 0,
                es_pro INTEGER DEFAULT 0,
                facturas_usadas INTEGER DEFAULT 0,
                fecha_registro TEXT,
                fecha_pro TEXT
            )
        """)
        
        # Migración segura: asegura que exista la columna email_verificado
        cursor.execute("PRAGMA table_info(usuarios)")
        columnas = [col[1] for col in cursor.fetchall()]
        if "email_verificado" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN email_verificado INTEGER DEFAULT 0")
            cursor.execute("UPDATE usuarios SET email_verificado = 1")
        if "tipo_licencia" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN tipo_licencia TEXT")
        if "stripe_customer_id" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN stripe_customer_id TEXT")
        if "stripe_subscription_id" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN stripe_subscription_id TEXT")
        if "subscription_status" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN subscription_status TEXT DEFAULT 'inactive'")
        if "subscription_period_end" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN subscription_period_end TEXT")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS verificaciones_email (
                email TEXT PRIMARY KEY,
                codigo TEXT NOT NULL,
                intentos INTEGER DEFAULT 0,
                fecha_expiracion TEXT NOT NULL
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS codigos_licencia (
                codigo TEXT PRIMARY KEY,
                usado INTEGER DEFAULT 0,
                usado_por_email TEXT,
                fecha_canje TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS transacciones_stripe (
                session_id TEXT PRIMARY KEY,
                email TEXT NOT NULL,
                monto_cents INTEGER,
                moneda TEXT,
                fecha_pago TEXT NOT NULL
            )
        """)
        
        # Carga opcional de clave maestra de pruebas exclusivamente desde st.secrets (Cero Hardcoding)
        clave_secrets = obtener_secreto("CLAVE_MAESTRA_TEST")
        if clave_secrets:
            cursor.execute("""
                INSERT OR IGNORE INTO codigos_licencia (codigo, usado)
                VALUES (?, 0)
            """, (clave_secrets.strip().upper(),))
        
        conn.commit()

        if managed_payments:
            try:
                managed_payments.initialize_managed_payments_schema()
            except Exception:
                pass


init_db()


# ==============================================================================
# VALIDACIÓN DE CORREO REAL Y ENVÍO SEGURO POR SMTP (TLS/SSL)
# ==============================================================================

def validar_email_real(email: str):
    """Valida sintaxis estricta, estructura de dominio y rechaza correos temporales."""
    email_clean = email.strip().lower()
    if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", email_clean):
        return False, "La dirección de correo electrónico no tiene un formato válido."
    
    partes = email_clean.split("@")
    if len(partes) != 2:
        return False, "Formato de correo inválido."
    
    dominio = partes[1]
    if dominio in DOMINIOS_DESECHABLES:
        return False, "Por motivos de seguridad fiscal, no se admiten correos temporales ni desechables."
    
    if "." not in dominio or len(dominio.split(".")[-1]) < 2:
        return False, "El dominio del correo electrónico no parece válido."
    
    return True, email_clean


def enviar_email_codigo(destinatario: str, codigo: str):
    """
    Envía un correo con diseño corporativo y código OTP de 6 dígitos mediante SMTP.
    Si no hay credenciales SMTP en secrets.toml, activa el modo simulación para desarrollo.
    """
    smtp_server = obtener_secreto("SMTP_SERVER")
    smtp_port = int(obtener_secreto("SMTP_PORT", "587"))
    smtp_user = obtener_secreto("SMTP_USER")
    smtp_pass = obtener_secreto("SMTP_PASSWORD")
    sender_name = obtener_secreto("SMTP_SENDER_NAME", "Facturae B2B Seguridad")
    sender_email = obtener_secreto("SMTP_SENDER_EMAIL", smtp_user)

    if not (smtp_server and smtp_user and smtp_pass):
        return False, "SMTP_NO_CONFIGURADO"
    
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"🔐 Tu código de verificación Facturae B2B: {codigo}"
        msg["From"] = f"{sender_name} <{sender_email}>"
        msg["To"] = destinatario

        html_body = f"""
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f8fafc; margin: 0; padding: 24px; color: #0f172a; }}
            .card {{ max-width: 500px; margin: 0 auto; background: #ffffff; border-radius: 16px; padding: 32px; border: 1px solid #e2e8f0; box-shadow: 0 10px 25px rgba(0,0,0,0.06); }}
            .header {{ text-align: center; border-bottom: 1px solid #f1f5f9; padding-bottom: 20px; margin-bottom: 24px; }}
            .header h2 {{ color: #2563eb; margin: 0; font-size: 24px; font-weight: 800; }}
            .badge {{ display: inline-block; background: #dbeafe; color: #1e40af; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 999px; text-transform: uppercase; margin-bottom: 12px; }}
            .code-box {{ background: #f0fdf4; border: 2px dashed #22c55e; border-radius: 12px; padding: 20px; text-align: center; font-size: 34px; font-weight: 800; letter-spacing: 8px; color: #15803d; margin: 24px 0; }}
            .info {{ font-size: 13px; color: #475569; line-height: 1.6; margin: 0 0 16px 0; }}
            .warning {{ font-size: 12px; color: #64748b; background: #f8fafc; border-left: 3px solid #cbd5e1; padding: 8px 12px; border-radius: 4px; }}
            .footer {{ text-align: center; margin-top: 30px; font-size: 11px; color: #94a3b8; }}
          </style>
        </head>
        <body>
          <div class="card">
            <div class="header">
              <span class="badge">Autenticación de Seguridad</span>
              <h2>Facturae B2B</h2>
            </div>
            <p>Hola,</p>
            <p class="info">Para verificar la autenticidad de tu correo electrónico y proteger tus datos fiscales confidenciales, introduce el siguiente código de seguridad en la aplicación:</p>
            <div class="code-box">{codigo}</div>
            <p class="info">⏰ Este código es de un solo uso y caduca en <b>15 minutos</b>.</p>
            <div class="warning">
              🔒 Si tú no has solicitado este código, puedes ignorar este mensaje; nadie puede activar la cuenta sin él.
            </div>
            <div class="footer">
              Facturae B2B · Plataforma Segura en Memoria Volátil
            </div>
          </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_body, "html", "utf-8"))

        context = ssl.create_default_context()
        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context, timeout=12) as server:
                server.login(smtp_user, smtp_pass)
                server.sendmail(sender_email, destinatario, msg.as_string())
        else:
            with smtplib.SMTP(smtp_server, smtp_port, timeout=12) as server:
                server.starttls(context=context)
                server.login(smtp_user, smtp_pass)
                server.sendmail(sender_email, destinatario, msg.as_string())
        
        return True, "Código enviado correctamente a tu buzón."
    except Exception as e:
        err_msg = str(e)
        return False, f"Error del servidor de correo: {err_msg}"


def generar_y_enviar_codigo_verificacion(email: str):
    """Genera un OTP criptográfico de 6 dígitos con validez de 15 minutos y lo envía."""
    codigo = f"{secrets.randbelow(900000) + 100000}"
    expira = (datetime.now() + timedelta(minutes=15)).strftime("%Y-%m-%d %H:%M:%S")
    
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO verificaciones_email (email, codigo, intentos, fecha_expiracion)
            VALUES (?, ?, 0, ?)
            ON CONFLICT(email) DO UPDATE SET
                codigo = excluded.codigo,
                intentos = 0,
                fecha_expiracion = excluded.fecha_expiracion
        """, (email.lower(), codigo, expira))
        conn.commit()

    ok_envio, msg_envio = enviar_email_codigo(email, codigo)
    return codigo, ok_envio, msg_envio


def validar_codigo_otp(email: str, codigo_ingresado: str):
    """Comprueba el código de 6 dígitos con protección contra ataques de fuerza bruta."""
    cod_clean = codigo_ingresado.strip()
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM verificaciones_email WHERE email = ?", (email.lower(),))
        registro = cursor.fetchone()
        
        if not registro:
            return False, "No hay ningún código pendiente para este correo. Solicita uno nuevo."
        
        if registro["fecha_expiracion"] < ahora:
            return False, "⏰ El código ha caducado (validez de 15 minutos). Solicita un nuevo código."
        
        if registro["intentos"] >= 5:
            return False, "❌ Has superado el número máximo de intentos fallidos. Solicita un nuevo código."
        
        if secrets.compare_digest(registro["codigo"], cod_clean):
            cursor.execute("UPDATE usuarios SET email_verificado = 1 WHERE email = ?", (email.lower(),))
            cursor.execute("DELETE FROM verificaciones_email WHERE email = ?", (email.lower(),))
            conn.commit()
            return True, "✅ Correo verificado con éxito."
        else:
            intentos_restantes = 4 - registro["intentos"]
            cursor.execute("UPDATE verificaciones_email SET intentos = intentos + 1 WHERE email = ?", (email.lower(),))
            conn.commit()
            return False, f"Código incorrecto. Te quedan {max(0, intentos_restantes)} intentos."


# ==============================================================================
# AUTENTICACIÓN Y CRIPTOGRAFÍA SEGURA DE CONTRASEÑAS (PBKDF2-HMAC)
# ==============================================================================

def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    pwd_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        100000
    ).hex()
    return f"{salt}:{pwd_hash}"


def verificar_password(password: str, stored_hash: str) -> bool:
    try:
        salt, expected_hash = stored_hash.split(":")
        actual_hash = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt.encode("utf-8"),
            100000
        ).hex()
        return secrets.compare_digest(actual_hash, expected_hash)
    except Exception:
        return False


def registrar_usuario(email: str, password: str):
    ok_email, msg_email = validar_email_real(email)
    if not ok_email:
        return False, msg_email
    
    if len(password) < 8 or not re.search(r"\d", password) or not re.search(r"[a-zA-Z]", password):
        return False, "Por confidencialidad fiscal, la contraseña debe tener al menos 8 caracteres y combinar letras y números."
    
    pwd_hashed = hash_password(password)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO usuarios (email, password_hash, email_verificado, es_pro, facturas_usadas, fecha_registro)
                VALUES (?, ?, 0, 0, 0, ?)
            """, (msg_email, pwd_hashed, ahora))
            conn.commit()
        
        codigo, ok_mail, detalle_mail = generar_y_enviar_codigo_verificacion(msg_email)
        return True, {
            "email": msg_email,
            "codigo": codigo,
            "enviado_smtp": ok_mail,
            "detalle_smtp": detalle_mail
        }
    except sqlite3.IntegrityError:
        return False, "Ya existe una cuenta registrada con este correo electrónico."
    except Exception as e:
        print(f"[ERROR registrar_usuario] {type(e).__name__}: {e}")
        return False, "Error interno al crear la cuenta. Por favor, inténtalo de nuevo más tarde."




def autenticar_usuario(email: str, password: str):
    email_clean = email.strip().lower()
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM usuarios WHERE email = ?", (email_clean,))
        user = cursor.fetchone()
        
        if not user or not verificar_password(password, user["password_hash"]):
            return False, None, "Correo electrónico o contraseña incorrectos."
        
        if not user["email_verificado"]:
            codigo, ok_mail, detalle_mail = generar_y_enviar_codigo_verificacion(email_clean)
            return False, {
                "necesita_verificar": True,
                "email": email_clean,
                "codigo": codigo,
                "enviado_smtp": ok_mail,
                "detalle_smtp": detalle_mail
            }, "Tu cuenta no ha sido verificada. Te hemos enviado un código a tu correo."
        
        user_keys = user.keys() if hasattr(user, "keys") else []
        return True, {
            "id": user["id"],
            "email": user["email"],
            "es_pro": bool(user["es_pro"]),
            "facturas_usadas": user["facturas_usadas"],
            "tipo_licencia": user["tipo_licencia"] if "tipo_licencia" in user_keys else None,
            "stripe_customer_id": user["stripe_customer_id"] if "stripe_customer_id" in user_keys else None,
            "stripe_subscription_id": user["stripe_subscription_id"] if "stripe_subscription_id" in user_keys else None,
            "subscription_status": user["subscription_status"] if "subscription_status" in user_keys else None,
        }, "Inicio de sesión correcto."


def recargar_usuario(email: str):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM usuarios WHERE email = ?", (email.lower(),))
        user = cursor.fetchone()
        if user:
            user_keys = user.keys() if hasattr(user, "keys") else []
            return {
                "id": user["id"],
                "email": user["email"],
                "es_pro": bool(user["es_pro"]),
                "facturas_usadas": user["facturas_usadas"],
                "tipo_licencia": user["tipo_licencia"] if "tipo_licencia" in user_keys else None,
                "stripe_customer_id": user["stripe_customer_id"] if "stripe_customer_id" in user_keys else None,
                "stripe_subscription_id": user["stripe_subscription_id"] if "stripe_subscription_id" in user_keys else None,
                "subscription_status": user["subscription_status"] if "subscription_status" in user_keys else None,
            }
    return None


def canjear_codigo_un_solo_uso(email: str, codigo: str):
    c = codigo.strip().upper()
    if not c:
        return False, "Por favor, introduce tu código de licencia."
    
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT codigo, usado, usado_por_email FROM codigos_licencia WHERE codigo = ?", (c,))
        row = cursor.fetchone()
        
        if not row:
            return False, "El código introducido no es válido o no existe."
        
        if row["usado"] == 1:
            return False, "❌ Este código ya ha sido canjeado anteriormente por otra persona y no puede reutilizarse."
        
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("""
            UPDATE codigos_licencia
            SET usado = 1, usado_por_email = ?, fecha_canje = ?
            WHERE codigo = ?
        """, (email.lower(), ahora, c))
        
        cursor.execute("""
            UPDATE usuarios
            SET es_pro = 1, fecha_pro = ?
            WHERE email = ?
        """, (ahora, email.lower()))
        
        conn.commit()
    
    return True, "🎉 ¡Código validado y canjeado con éxito! Tu cuenta ahora es PRO ilimitada."


def registrar_factura_generada():
    if st.session_state.get("usuario_actual"):
        email = st.session_state.usuario_actual["email"]
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE usuarios SET facturas_usadas = facturas_usadas + 1 WHERE email = ?", (email,))
            conn.commit()
        st.session_state.usuario_actual = recargar_usuario(email)


# ==============================================================================
# INTEGRACIÓN DE PASARELA STRIPE CHECKOUT Y ACTUALIZACIÓN AUTOMÁTICA A PRO
# ==============================================================================

def crear_sesion_pago_stripe(email: Optional[str] = None, base_url: str = "http://localhost:8501", tipo_plan: str = "vitalicia"):
    """
    Crea una sesión de Stripe Checkout para los 2 planes oficiales:
    - 'vitalicia': Pago único de 99,49 € (mode="payment")
    - 'mensual': Suscripción recurrente de 9,49 € / mes (mode="subscription")
    """
    if not (stripe and STRIPE_SECRET_KEY):
        return None, "Stripe API no está configurado."
    
    stripe.api_key = STRIPE_SECRET_KEY
    clean_base_url = (base_url or "http://localhost:8501").rstrip("/")
    success_url = f"{clean_base_url}/?session_id={{CHECKOUT_SESSION_ID}}&plan={tipo_plan}"
    cancel_url = f"{clean_base_url}/?pago_cancelado=1"

    email_clean = (email or "").strip().lower() or None

    try:
        if tipo_plan == "mensual":
            price_id = STRIPE_PRICE_MENSUAL or "price_1USGYQROobv6Wd8mr2jwehcN"
            metadata = {
                "tipo_plan": "mensual",
                "producto": "FacturaeB2B_Suscripcion_Mensual"
            }
            if email_clean:
                metadata["email"] = email_clean

            session_params = {
                "line_items": [{"price": price_id, "quantity": 1}],
                "mode": "subscription",
                "tax_id_collection": {"enabled": True},
                "success_url": success_url,
                "cancel_url": cancel_url,
                "metadata": metadata,
                "subscription_data": {
                    "metadata": metadata
                }
            }
            if email_clean:
                session_params["customer_email"] = email_clean
                session_params["client_reference_id"] = email_clean

            session = stripe.checkout.Session.create(**session_params)
        else:
            price_id = STRIPE_PRICE_VITALICIA or "price_1USGYQROobv6Wd8mcdXDWFl1"
            metadata = {
                "tipo_plan": "vitalicia",
                "producto": "FacturaeB2B_Licencia_Vitalicia"
            }
            if email_clean:
                metadata["email"] = email_clean

            session_params = {
                "line_items": [{"price": price_id, "quantity": 1}],
                "mode": "payment",
                "tax_id_collection": {"enabled": True},
                "success_url": success_url,
                "cancel_url": cancel_url,
                "metadata": metadata
            }
            if email_clean:
                session_params["customer_email"] = email_clean
                session_params["client_reference_id"] = email_clean

            session = stripe.checkout.Session.create(**session_params)
        return session.url, None
    except Exception as e:
        return None, f"Error al generar la pasarela de pago: {str(e)}"



def verificar_y_activar_sesion_stripe(session_id: str, email_usuario_actual: str):
    """
    Verifica con la API de Stripe en backend que el pago se completó
    y activa la cuenta a nivel PRO distinguiendo el tipo de plan:
    - 'vitalicia' (99,49 € pago único de por vida)
    - 'mensual' (9,49 €/mes activo mientras se mantenga la suscripción)
    """
    s_id = session_id.strip()
    if not s_id:
        return False, "Identificador de sesión de pago no válido."
    
    # 1. Comprobar si esta sesión ya fue canjeada en SQLite (Anti-Replay)
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT session_id, email FROM transacciones_stripe WHERE session_id = ?", (s_id,))
        if cursor.fetchone():
            return False, "Esta sesión de pago ya ha sido procesada anteriormente."
    
    # 2. Modo desarrollo simulado (solo si MODO_DESARROLLO=True y session_id empieza por MOCK-PAY-)
    if MODO_DESARROLLO and s_id.startswith("MOCK-PAY-"):
        tipo_sim = "mensual" if "MENSUAL" in s_id else "vitalicia"
        monto_sim = 949 if tipo_sim == "mensual" else 9949
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR IGNORE INTO transacciones_stripe (session_id, email, monto_cents, moneda, fecha_pago)
                VALUES (?, ?, ?, 'eur', ?)
            """, (s_id, email_usuario_actual.lower(), monto_sim, ahora))
            cursor.execute("""
                UPDATE usuarios
                SET es_pro = 1,
                    fecha_pro = ?,
                    tipo_licencia = ?,
                    subscription_status = 'active'
                WHERE email = ?
            """, (ahora, tipo_sim, email_usuario_actual.lower()))
            conn.commit()
        msg_sim = (
            "🎉 ¡Suscripción mensual de 9,49 € verificada! Tu cuenta PRO está activa."
            if tipo_sim == "mensual"
            else "🎉 ¡Pago de 99,49 € verificado! Tu cuenta ahora es PRO ilimitada de por vida."
        )
        return True, msg_sim

    # 3. Consulta estricta y segura a la API de Stripe con clave secreta
    if not (stripe and STRIPE_SECRET_KEY):
        return False, "La API de Stripe no está configurada en el servidor."
    
    try:
        stripe.api_key = STRIPE_SECRET_KEY
        session = stripe.checkout.Session.retrieve(s_id)
        
        # Validar estado del pago
        if session.payment_status not in ("paid", "no_payment_required") and getattr(session, "status", "") != "complete":
            return False, "El pago en Stripe no se ha completado correctamente."
        
        # Validar que la sesión corresponde al usuario conectado
        email_sesion = (session.client_reference_id or (session.customer_details and session.customer_details.email) or "").strip().lower()
        if email_sesion and email_sesion != email_usuario_actual.strip().lower():
            return False, "Esta sesión de pago pertenece a otra cuenta de usuario."
        
        # Determinar tipo de licencia: mensual o vitalicia
        es_suscripcion = (session.mode == "subscription") or bool(getattr(session, "subscription", None))
        tipo_licencia = "mensual" if es_suscripcion else "vitalicia"
        cust_id = getattr(session, "customer", None)
        sub_id = getattr(session, "subscription", None)
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO transacciones_stripe
                (session_id, email, monto_cents, moneda, fecha_pago, customer_id, subscription_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (s_id, email_usuario_actual.lower(), session.amount_total, str(session.currency).lower(), ahora, cust_id, sub_id))

            cursor.execute("""
                UPDATE usuarios
                SET es_pro = 1,
                    fecha_pro = COALESCE(fecha_pro, ?),
                    tipo_licencia = ?,
                    stripe_customer_id = COALESCE(?, stripe_customer_id),
                    stripe_subscription_id = COALESCE(?, stripe_subscription_id),
                    subscription_status = 'active'
                WHERE email = ?
            """, (ahora, tipo_licencia, cust_id, sub_id, email_usuario_actual.lower()))
            conn.commit()
            
        if tipo_licencia == "vitalicia":
            return True, "🎉 ¡Pago de 99,49 € verificado con éxito! Tu cuenta ahora es PRO vitalicia de por vida."
        else:
            return True, "🎉 ¡Suscripción mensual de 9,49 € activada con éxito! Tu cuenta ahora es PRO activa."
    except Exception as e:
        print(f"[ERROR verificar_y_activar_sesion_stripe] {type(e).__name__}: {e}")
        return False, "No se pudo verificar el pago con Stripe en este momento. Si se ha completado el cargo, tu cuenta se activará automáticamente vía Webhook."



def render_planes_pago_stripe(email: Optional[str] = None, key_prefix: str = "sb", stack: bool = False):
    """
    Renderiza en la web las 2 opciones oficiales de Stripe:
    1. Suscripción Mensual (9,49 € / mes)
    2. Licencia PRO Vitalicia (99,49 € pago único)
    """
    email_clean = (email or "").strip().lower()

    # Preparamos las URLs dinámicas de Stripe Checkout
    url_m, err_m = (None, None)
    url_v, err_v = (None, None)
    if stripe and STRIPE_SECRET_KEY:
        url_m, err_m = crear_sesion_pago_stripe(email_clean, APP_BASE_URL, tipo_plan="mensual")
        url_v, err_v = crear_sesion_pago_stripe(email_clean, APP_BASE_URL, tipo_plan="vitalicia")

    if stack:
        container_m = st.container(border=True)
        container_v = st.container(border=True)
    else:
        col_m, col_v = st.columns(2)
        container_m = col_m.container(border=True)
        container_v = col_v.container(border=True)

    with container_m:
        st.markdown("##### 💳 Suscripción Mensual")
        st.markdown("### **9,49 €** <span style='font-size:14px; color:#64748b;'>/ mes</span>", unsafe_allow_html=True)
        st.caption("✨ Facturas y descargas ilimitadas. Cancela cuando quieras.")
        if url_m:
            st.link_button("💳 Suscribirme (9,49 €/mes)", url_m, type="secondary", use_container_width=True)
        elif STRIPE_PAYMENT_URL:
            st.link_button("💳 Suscribirme (9,49 €/mes)", STRIPE_PAYMENT_URL, use_container_width=True)
        else:
            st.error("Error al cargar pasarela mensual.")

    with container_v:
        st.markdown("##### 👑 Licencia Vitalicia")
        st.markdown("### **99,49 €** <span style='font-size:14px; color:#16a34a;'>pago único</span>", unsafe_allow_html=True)
        st.caption("🚀 Acceso PRO de por vida. Nunca más volverás a pagar.")
        if url_v:
            st.link_button("👑 Comprar Vitalicia (99,49 €)", url_v, type="primary", use_container_width=True)
        elif STRIPE_PAYMENT_URL:
            st.link_button("👑 Comprar Vitalicia (99,49 €)", STRIPE_PAYMENT_URL, use_container_width=True)
        else:
            st.error("Error al cargar pasarela vitalicia.")

    st.caption("🔒 Pago 100% seguro procesado por Stripe (Tarjeta, Apple Pay, Google Pay). Factura oficial con NIF/CIF.")

    if MODO_DESARROLLO:
        st.markdown("---")
        st.caption("🧪 **Simulador de pagos inmediatos (Modo Desarrollo)**")
        c_sim1, c_sim2 = st.columns(2)
        with c_sim1:
            if st.button("🧪 Simular Mensual (9,49 €)", key=f"btn_sim_m_{key_prefix}", use_container_width=True):
                if not email_clean:
                    st.warning("⚠️ Inicia sesión o regístrate en la barra lateral para simular la activación.")
                else:
                    sim_id = f"MOCK-PAY-MENSUAL-{secrets.token_hex(6)}"
                    ok_s, msg_s = verificar_y_activar_sesion_stripe(sim_id, email_clean)
                    if ok_s:
                        st.session_state.usuario_actual = recargar_usuario(email_clean)
                        st.balloons()
                        st.success(msg_s)
                        st.rerun()
        with c_sim2:
            if st.button("🧪 Simular Vitalicia (99,49 €)", key=f"btn_sim_v_{key_prefix}", use_container_width=True):
                if not email_clean:
                    st.warning("⚠️ Inicia sesión o regístrate en la barra lateral para simular la activación.")
                else:
                    sim_id = f"MOCK-PAY-VITALICIA-{secrets.token_hex(6)}"
                    ok_s, msg_s = verificar_y_activar_sesion_stripe(sim_id, email_clean)
                    if ok_s:
                        st.session_state.usuario_actual = recargar_usuario(email_clean)
                        st.balloons()
                        st.success(msg_s)
                        st.rerun()



def render_boton_checkout_stripe(email: str, label: str = "💳 Comprar Licencia PRO", key_prefix: str = "sb"):
    """Mantiene compatibilidad redirigiendo al selector oficial de planes."""
    render_planes_pago_stripe(email, key_prefix=key_prefix)


# ==============================================================================
# ESTADO DE SESIÓN
# ==============================================================================

if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = None

if "pendiente_verificacion" not in st.session_state:
    st.session_state.pendiente_verificacion = None


# ==============================================================================
# FUNCIONES DE NORMALIZACIÓN Y SANITIZACIÓN
# ==============================================================================

def fmt_dec(val):
    return f"{float(val):.2f}"

def limpiar_texto(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0") and re.match(r"^-?\d+\.0$", s):
        s = s[:-2]
    return s

def parsear_float_seguro(val, default=0.0):
    if pd.isna(val):
        return default
    s = str(val).strip().replace("€", "").replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return default

def normalizar_cp(val):
    s = limpiar_texto(val)
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"\D", "", s)
    return s.zfill(5)[:5] if s else ""

def normalizar_fecha(val):
    if pd.isna(val) or not str(val).strip():
        return str(date.today())
    s_val = str(val).strip()
    # Intentar formato estándar YYYY-MM-DD primero para evitar warnings
    try:
        return datetime.strptime(s_val[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
    except Exception:
        pass
    try:
        parsed = pd.to_datetime(val, dayfirst=True)
        return parsed.strftime("%Y-%m-%d")
    except Exception:
        return s_val[:10]

def sanitizar_csv_field(val: str) -> str:
    s = str(val).strip()
    if s and s[0] in ("=", "+", "-", "@", "\t", "\r"):
        s = "'" + s
    return s[:2000]

def guardar_reporte_error(categoria, descripcion):
    try:
        escribir_cabecera = (not LOG_ERRORES_PATH.exists()) or LOG_ERRORES_PATH.stat().st_size == 0
        with LOG_ERRORES_PATH.open("a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNAS_LOG_ERRORES)
            if escribir_cabecera:
                writer.writeheader()
            writer.writerow({
                "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Categoria": sanitizar_csv_field(categoria),
                "Descripcion_Usuario": sanitizar_csv_field(descripcion),
            })
        return True, "Mensaje enviado con éxito. Nuestro equipo revisará tu consulta lo antes posible."
    except Exception:
        return False, "No se pudo enviar la consulta en este momento. Por favor, inténtalo de nuevo más tarde."

def validar_datos(df_cabecera, df_lineas):
    """
    Valida cabecera y líneas del Excel, detectando inteligentemente las columnas
    tanto si vienen con el encabezado de diseño como simplificadas.
    """
    errores = []
    
    # Detección inteligente y flexible de columnas (resuelve discrepancias con 'Campo' vs 'Campo (No modificar)')
    col_campo = next((c for c in df_cabecera.columns if str(c).strip().lower().startswith("campo")), None)
    col_valor = next((c for c in df_cabecera.columns if str(c).strip().lower().startswith("valor")), None)

    if not col_campo or not col_valor:
        return ["La plantilla no contiene las columnas requeridas 'Campo' y 'Valor'."], {}

    raw_dict = df_cabecera.set_index(col_campo)[col_valor].to_dict()
    datos = {limpiar_texto(k): v for k, v in raw_dict.items() if pd.notna(k)}
    
    campos_obligatorios = [
        "Serie_Factura", "Numero_Factura", "Fecha_Emision", "IBAN_Pago",
        "Emisor_Tipo", "Emisor_NIF", "Emisor_Direccion", "Emisor_CP", "Emisor_Ciudad", "Emisor_Provincia",
        "Receptor_Tipo", "Receptor_NIF", "Receptor_Direccion", "Receptor_CP", "Receptor_Ciudad", "Receptor_Provincia"
    ]
    
    for campo in campos_obligatorios:
        if not limpiar_texto(datos.get(campo, "")):
            errores.append(f"Falta rellenar la casilla obligatoria: '{campo}'.")

    emisor_tipo = limpiar_texto(datos.get("Emisor_Tipo")).upper()
    if emisor_tipo == "F":
        if not limpiar_texto(datos.get("Emisor_Nombre")) or not limpiar_texto(datos.get("Emisor_Apellido1")):
            errores.append("Si el Emisor es Autónomo ('F'), falta rellenar 'Emisor_Nombre' y 'Emisor_Apellido1'.")
    elif emisor_tipo == "J":
        if not limpiar_texto(datos.get("Emisor_Razon_Social")):
            errores.append("Si el Emisor es Empresa ('J'), falta rellenar 'Emisor_Razon_Social'.")
    else:
        errores.append("La casilla 'Emisor_Tipo' debe ser 'F' (Autónomo) o 'J' (Empresa).")

    receptor_tipo = limpiar_texto(datos.get("Receptor_Tipo")).upper()
    if receptor_tipo == "F":
        if not limpiar_texto(datos.get("Receptor_Nombre")) or not limpiar_texto(datos.get("Receptor_Apellido1")):
            errores.append("Si el Cliente es Persona Física ('F'), falta rellenar 'Receptor_Nombre' y 'Receptor_Apellido1'.")
    elif receptor_tipo == "J":
        if not limpiar_texto(datos.get("Receptor_Razon_Social")):
            errores.append("Si el Cliente es Empresa o Ayuntamiento ('J'), falta rellenar 'Receptor_Razon_Social'.")
    else:
        errores.append("La casilla 'Receptor_Tipo' debe ser 'F' (Persona Física) o 'J' (Empresa/Organismo).")

    iban = re.sub(r"\s+", "", limpiar_texto(datos.get("IBAN_Pago", ""))).upper()
    if not re.match(r"^ES\d{22}$", iban):
        errores.append("La casilla 'IBAN_Pago' debe tener formato español válido (ES + 22 números continuos).")

    lineas_validas = 0
    col_concepto = next((c for c in df_lineas.columns if str(c).strip().lower().startswith("concepto")), "Concepto")
    if not df_lineas.empty and col_concepto in df_lineas.columns:
        for _, row in df_lineas.iterrows():
            if limpiar_texto(row.get(col_concepto)):
                lineas_validas += 1
    if lineas_validas == 0:
        errores.append("La pestaña 'Lineas_Factura' no contiene ningún concepto facturado.")

    return errores, datos


# ==============================================================================
# MOTOR XML FACTURAE 3.2.2 (100% PROCESADO EN MEMORIA VOLÁTIL)
# ==============================================================================

def construir_xml_facturae(datos, df_lineas):
    ns = "http://www.facturae.gob.es/formato/Versiones/Facturaev3_2_2.xml"
    ET.register_namespace("fe", ns)
    root = ET.Element(f"{{{ns}}}Facturae")

    base_imponible_total = 0.0
    iva_porcentaje = parsear_float_seguro(datos.get("IVA_Porcentaje", 21.0), 21.0)
    irpf_porcentaje = parsear_float_seguro(datos.get("IRPF_Porcentaje", 0.0), 0.0)

    col_concepto = next((c for c in df_lineas.columns if str(c).strip().lower().startswith("concepto")), "Concepto")
    col_cantidad = next((c for c in df_lineas.columns if str(c).strip().lower().startswith("cantidad")), "Cantidad")
    col_precio = next((c for c in df_lineas.columns if str(c).strip().lower().startswith("precio")), "Precio_Unitario")

    lineas_procesadas = []
    for _, row in df_lineas.iterrows():
        desc = limpiar_texto(row.get(col_concepto))
        if not desc:
            continue
        cant = parsear_float_seguro(row.get(col_cantidad, 1), 1.0)
        precio = parsear_float_seguro(row.get(col_precio, 0), 0.0)
        importe_linea = round(cant * precio, 2)
        base_imponible_total += importe_linea
        lineas_procesadas.append({
            "Concepto": desc,
            "Cantidad": cant,
            "Precio_Unitario": precio,
            "Subtotal": importe_linea
        })

    base_imponible_total = round(base_imponible_total, 2)
    cuota_iva = round(base_imponible_total * (iva_porcentaje / 100.0), 2)
    cuota_irpf = round(base_imponible_total * (irpf_porcentaje / 100.0), 2)
    total_factura = round(base_imponible_total + cuota_iva - cuota_irpf, 2)

    serie = limpiar_texto(datos.get("Serie_Factura"))
    num_fac = limpiar_texto(datos.get("Numero_Factura"))
    emisor_nif = limpiar_texto(datos.get("Emisor_NIF")).upper()
    batch_id = f"{emisor_nif}{num_fac}{serie}"
    fecha_emision = normalizar_fecha(datos.get("Fecha_Emision"))

    fh = ET.SubElement(root, "FileHeader")
    ET.SubElement(fh, "SchemaVersion").text = "3.2.2"
    ET.SubElement(fh, "Modality").text = "I"
    ET.SubElement(fh, "InvoiceIssuerType").text = "EM"
    batch = ET.SubElement(fh, "Batch")
    ET.SubElement(batch, "BatchIdentifier").text = batch_id
    ET.SubElement(batch, "InvoicesCount").text = "1"
    
    for tag_name in ["TotalInvoicesAmount", "TotalOutstandingAmount", "TotalExecutableAmount"]:
        node = ET.SubElement(batch, tag_name)
        ET.SubElement(node, "TotalAmount").text = fmt_dec(total_factura)
    ET.SubElement(batch, "InvoiceCurrencyCode").text = "EUR"

    parties = ET.SubElement(root, "Parties")
    
    def add_party(parent_tag, prefix, include_dir3=False):
        party = ET.SubElement(parties, parent_tag)
        tax_id = ET.SubElement(party, "TaxIdentification")
        tipo = limpiar_texto(datos.get(f"{prefix}_Tipo")).upper()
        ET.SubElement(tax_id, "PersonTypeCode").text = tipo
        ET.SubElement(tax_id, "ResidenceTypeCode").text = "R"
        ET.SubElement(tax_id, "TaxIdentificationNumber").text = limpiar_texto(datos.get(f"{prefix}_NIF")).upper()

        if include_dir3 and limpiar_texto(datos.get("DIR3_Oficina_Contable")):
            centres = ET.SubElement(party, "AdministrativeCentres")
            roles = [
                ("01", "DIR3_Oficina_Contable", "Oficina Contable"),
                ("02", "DIR3_Organo_Gestor", "Organo Gestor"),
                ("03", "DIR3_Unidad_Tramitadora", "Unidad Tramitadora")
            ]
            for role_code, field_key, desc_role in roles:
                code_val = limpiar_texto(datos.get(field_key))
                if code_val:
                    ac = ET.SubElement(centres, "AdministrativeCentre")
                    ET.SubElement(ac, "CentreCode").text = code_val
                    ET.SubElement(ac, "RoleTypeCode").text = role_code
                    addr_ac = ET.SubElement(ac, "AddressInSpain")
                    ET.SubElement(addr_ac, "Address").text = limpiar_texto(datos.get(f"{prefix}_Direccion"))
                    ET.SubElement(addr_ac, "PostCode").text = normalizar_cp(datos.get(f"{prefix}_CP"))
                    ET.SubElement(addr_ac, "Town").text = limpiar_texto(datos.get(f"{prefix}_Ciudad"))
                    ET.SubElement(addr_ac, "Province").text = limpiar_texto(datos.get(f"{prefix}_Provincia"))
                    ET.SubElement(addr_ac, "CountryCode").text = "ESP"
                    ET.SubElement(ac, "CentreDescription").text = desc_role

        if tipo == "F":
            indiv = ET.SubElement(party, "Individual")
            ET.SubElement(indiv, "Name").text = limpiar_texto(datos.get(f"{prefix}_Nombre"))
            ET.SubElement(indiv, "FirstSurname").text = limpiar_texto(datos.get(f"{prefix}_Apellido1"))
            ap2 = limpiar_texto(datos.get(f"{prefix}_Apellido2"))
            if ap2:
                ET.SubElement(indiv, "SecondSurname").text = ap2
            addr = ET.SubElement(indiv, "AddressInSpain")
        else:
            legal = ET.SubElement(party, "LegalEntity")
            ET.SubElement(legal, "CorporateName").text = limpiar_texto(datos.get(f"{prefix}_Razon_Social"))
            addr = ET.SubElement(legal, "AddressInSpain")

        ET.SubElement(addr, "Address").text = limpiar_texto(datos.get(f"{prefix}_Direccion"))
        ET.SubElement(addr, "PostCode").text = normalizar_cp(datos.get(f"{prefix}_CP"))
        ET.SubElement(addr, "Town").text = limpiar_texto(datos.get(f"{prefix}_Ciudad"))
        ET.SubElement(addr, "Province").text = limpiar_texto(datos.get(f"{prefix}_Provincia"))
        ET.SubElement(addr, "CountryCode").text = "ESP"

    add_party("SellerParty", "Emisor", include_dir3=False)
    add_party("BuyerParty", "Receptor", include_dir3=True)

    invoices = ET.SubElement(root, "Invoices")
    inv = ET.SubElement(invoices, "Invoice")
    
    inv_header = ET.SubElement(inv, "InvoiceHeader")
    ET.SubElement(inv_header, "InvoiceNumber").text = num_fac
    ET.SubElement(inv_header, "InvoiceSeriesCode").text = serie
    ET.SubElement(inv_header, "InvoiceDocumentType").text = "FC"
    ET.SubElement(inv_header, "InvoiceClass").text = "OO"

    issue_data = ET.SubElement(inv, "InvoiceIssueData")
    ET.SubElement(issue_data, "IssueDate").text = fecha_emision
    ET.SubElement(issue_data, "InvoiceCurrencyCode").text = "EUR"
    ET.SubElement(issue_data, "TaxCurrencyCode").text = "EUR"
    ET.SubElement(issue_data, "LanguageName").text = "es"

    taxes_out = ET.SubElement(inv, "TaxesOutputs")
    tax_iva = ET.SubElement(taxes_out, "Tax")
    ET.SubElement(tax_iva, "TaxTypeCode").text = "01"
    ET.SubElement(tax_iva, "TaxRate").text = fmt_dec(iva_porcentaje)
    tb_iva = ET.SubElement(tax_iva, "TaxableBase")
    ET.SubElement(tb_iva, "TotalAmount").text = fmt_dec(base_imponible_total)
    ta_iva = ET.SubElement(tax_iva, "TaxAmount")
    ET.SubElement(ta_iva, "TotalAmount").text = fmt_dec(cuota_iva)

    if irpf_porcentaje > 0:
        taxes_withheld = ET.SubElement(inv, "TaxesWithheld")
        tax_irpf = ET.SubElement(taxes_withheld, "Tax")
        ET.SubElement(tax_irpf, "TaxTypeCode").text = "04"
        ET.SubElement(tax_irpf, "TaxRate").text = fmt_dec(irpf_porcentaje)
        tb_irpf = ET.SubElement(tax_irpf, "TaxableBase")
        ET.SubElement(tb_irpf, "TotalAmount").text = fmt_dec(base_imponible_total)
        ta_irpf = ET.SubElement(tax_irpf, "TaxAmount")
        ET.SubElement(ta_irpf, "TotalAmount").text = fmt_dec(cuota_irpf)

    totals = ET.SubElement(inv, "InvoiceTotals")
    ET.SubElement(totals, "TotalGrossAmount").text = fmt_dec(base_imponible_total)
    ET.SubElement(totals, "TotalGrossAmountBeforeTaxes").text = fmt_dec(base_imponible_total)
    ET.SubElement(totals, "TotalTaxOutputs").text = fmt_dec(cuota_iva)
    ET.SubElement(totals, "TotalTaxesWithheld").text = fmt_dec(cuota_irpf)
    ET.SubElement(totals, "InvoiceTotal").text = fmt_dec(total_factura)
    ET.SubElement(totals, "TotalOutstandingAmount").text = fmt_dec(total_factura)
    ET.SubElement(totals, "TotalExecutableAmount").text = fmt_dec(total_factura)

    items = ET.SubElement(inv, "Items")
    for linea in lineas_procesadas:
        il = ET.SubElement(items, "InvoiceLine")
        ET.SubElement(il, "ItemDescription").text = linea["Concepto"]
        ET.SubElement(il, "Quantity").text = fmt_dec(linea["Cantidad"])
        ET.SubElement(il, "UnitOfMeasure").text = "01"
        ET.SubElement(il, "UnitPriceWithoutTax").text = fmt_dec(linea["Precio_Unitario"])
        ET.SubElement(il, "TotalCost").text = fmt_dec(linea["Subtotal"])
        ET.SubElement(il, "GrossAmount").text = fmt_dec(linea["Subtotal"])
        
        l_taxes_out = ET.SubElement(il, "TaxesOutputs")
        l_tax_iva = ET.SubElement(l_taxes_out, "Tax")
        ET.SubElement(l_tax_iva, "TaxTypeCode").text = "01"
        ET.SubElement(l_tax_iva, "TaxRate").text = fmt_dec(iva_porcentaje)
        l_tb = ET.SubElement(l_tax_iva, "TaxableBase")
        ET.SubElement(l_tb, "TotalAmount").text = fmt_dec(linea["Subtotal"])
        l_ta = ET.SubElement(l_tax_iva, "TaxAmount")
        ET.SubElement(l_ta, "TotalAmount").text = fmt_dec(round(linea["Subtotal"] * iva_porcentaje / 100.0, 2))

    payment = ET.SubElement(inv, "PaymentDetails")
    inst = ET.SubElement(payment, "Installment")
    ET.SubElement(inst, "InstallmentDueDate").text = fecha_emision
    ET.SubElement(inst, "InstallmentAmount").text = fmt_dec(total_factura)
    ET.SubElement(inst, "PaymentMeans").text = "04"
    acc = ET.SubElement(inst, "AccountToBeCredited")
    ET.SubElement(acc, "IBAN").text = re.sub(r"\s+", "", limpiar_texto(datos.get("IBAN_Pago"))).upper()

    xml_str = ET.tostring(root, encoding="utf-8")
    pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ", encoding="UTF-8")
    return pretty_xml, base_imponible_total, cuota_iva, cuota_irpf, total_factura, pd.DataFrame(lineas_procesadas)


# ==============================================================================
# PLANTILLA EXCEL GENERADA EN MEMORIA
# ==============================================================================

def generar_plantilla_excel():
    output = BytesIO()
    df_cabecera = pd.DataFrame([
        {"Campo": "Serie_Factura", "Valor": "2026", "Explicacion": "Serie de tu factura (ej. 2026, A, FAC)"},
        {"Campo": "Numero_Factura", "Valor": "001", "Explicacion": "Número correlativo de la factura (ej. 001, 002)"},
        {"Campo": "Fecha_Emision", "Valor": str(date.today()), "Explicacion": "Fecha de emisión en formato AAAA-MM-DD"},
        {"Campo": "IVA_Porcentaje", "Valor": 21, "Explicacion": "Tipo de IVA aplicable: elige 21, 10, 4 o 0"},
        {"Campo": "IRPF_Porcentaje", "Valor": 0, "Explicacion": "Retención de IRPF (autónomos profesionales): 15, 7 o 0"},
        {"Campo": "IBAN_Pago", "Valor": "ES0000000000000000000000", "Explicacion": "Cuenta IBAN completa (ES + 22 dígitos sin espacios)"},
        {"Campo": "Emisor_Tipo", "Valor": "J", "Explicacion": "Escribe F si eres Autónomo o J si eres Empresa (SL/SA)"},
        {"Campo": "Emisor_NIF", "Valor": "B12345678", "Explicacion": "CIF o NIF del emisor (ejemplo ficticio)"},
        {"Campo": "Emisor_Nombre", "Valor": "", "Explicacion": "Nombre de pila (Solo si Emisor_Tipo es F)"},
        {"Campo": "Emisor_Apellido1", "Valor": "", "Explicacion": "Primer apellido (Solo si Emisor_Tipo es F)"},
        {"Campo": "Emisor_Apellido2", "Valor": "", "Explicacion": "Segundo apellido (Solo si Emisor_Tipo es F)"},
        {"Campo": "Emisor_Razon_Social", "Valor": "EMPRESA EJEMPLO S.L.", "Explicacion": "Razón social del emisor (Si Emisor_Tipo es J)"},
        {"Campo": "Emisor_Direccion", "Valor": "CALLE EJEMPLO 123", "Explicacion": "Domicilio fiscal ficticio"},
        {"Campo": "Emisor_CP", "Valor": "28001", "Explicacion": "Código postal de 5 dígitos"},
        {"Campo": "Emisor_Ciudad", "Valor": "MADRID", "Explicacion": "Municipio fiscal"},
        {"Campo": "Emisor_Provincia", "Valor": "MADRID", "Explicacion": "Provincia fiscal"},
        {"Campo": "Receptor_Tipo", "Valor": "J", "Explicacion": "F = Particular/Autónomo | J = Empresa o Administración"},
        {"Campo": "Receptor_NIF", "Valor": "B98765432", "Explicacion": "CIF o NIF del receptor"},
        {"Campo": "Receptor_Nombre", "Valor": "", "Explicacion": "Nombre del cliente (Solo si Receptor_Tipo es F)"},
        {"Campo": "Receptor_Apellido1", "Valor": "", "Explicacion": "Primer apellido del cliente (Solo si Receptor_Tipo es F)"},
        {"Campo": "Receptor_Apellido2", "Valor": "", "Explicacion": "Segundo apellido del cliente (Solo si Receptor_Tipo es F)"},
        {"Campo": "Receptor_Razon_Social", "Valor": "CLIENTE CORPORATIVO S.L.", "Explicacion": "Razón Social o Ayuntamiento receptor (Si es J)"},
        {"Campo": "Receptor_Direccion", "Valor": "AVENIDA PRINCIPAL 45", "Explicacion": "Dirección fiscal del cliente"},
        {"Campo": "Receptor_CP", "Valor": "28002", "Explicacion": "Código postal del cliente"},
        {"Campo": "Receptor_Ciudad", "Valor": "MADRID", "Explicacion": "Municipio del cliente"},
        {"Campo": "Receptor_Provincia", "Valor": "MADRID", "Explicacion": "Provincia del cliente"},
        {"Campo": "DIR3_Oficina_Contable", "Valor": "", "Explicacion": "Opcional (Solo FACe / Administración Pública): ej. L01280001"},
        {"Campo": "DIR3_Organo_Gestor", "Valor": "", "Explicacion": "Opcional (Solo FACe / Administración Pública): Código Órgano Gestor"},
        {"Campo": "DIR3_Unidad_Tramitadora", "Valor": "", "Explicacion": "Opcional (Solo FACe / Administración Pública): Código Unidad Tramitadora"},
    ])

    df_lineas = pd.DataFrame([
        {"Concepto": "Servicios de consultoría y arquitectura técnica de software", "Cantidad": 1, "Precio_Unitario": 500.00, "Subtotal_Visual": "=B2*C2"},
        {"Concepto": "Mantenimiento mensual y auditoría de ciberseguridad", "Cantidad": 2, "Precio_Unitario": 45.50, "Subtotal_Visual": "=B3*C3"}
    ])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_cabecera.to_excel(writer, sheet_name="Datos_Factura", index=False)
        df_lineas.to_excel(writer, sheet_name="Lineas_Factura", index=False)
        
        wb = writer.book
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        font_campo = Font(name="Calibri", size=11, bold=True, color="1E293B")
        fill_campo = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        font_valor = Font(name="Calibri", size=11, bold=True, color="0369A1")
        fill_valor = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        fill_valor_obligatorio = PatternFill(start_color="F0F9FF", end_color="F0F9FF", fill_type="solid")
        font_explicacion = Font(name="Calibri", size=10, italic=True, color="475569")
        fill_explicacion = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        
        thin_side = Side(border_style="thin", color="CBD5E1")
        med_blue = Side(border_style="medium", color="0284C7")
        border_standard = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        border_input = Border(left=med_blue, right=thin_side, top=thin_side, bottom=thin_side)

        ws1 = wb["Datos_Factura"]
        ws1.views.sheetView[0].showGridLines = True
        ws1.freeze_panes = "A2"
        ws1.column_dimensions["A"].width = 28
        ws1.column_dimensions["B"].width = 36
        ws1.column_dimensions["C"].width = 65
        ws1.row_dimensions[1].height = 30

        headers_ws1 = ["Campo (No modificar)", "Valor (Rellenar aquí)", "Guía rápida e instrucciones de formato"]
        for col_idx, text in enumerate(headers_ws1, 1):
            cell = ws1.cell(row=1, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="left" if col_idx != 2 else "center", vertical="center")

        for row_idx in range(2, len(df_cabecera) + 2):
            ws1.row_dimensions[row_idx].height = 25
            c_campo = ws1.cell(row=row_idx, column=1)
            c_valor = ws1.cell(row=row_idx, column=2)
            c_exp = ws1.cell(row=row_idx, column=3)
            
            c_campo.font = font_campo
            c_campo.fill = fill_campo
            c_campo.border = border_standard
            c_campo.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            
            c_valor.font = font_valor
            c_valor.fill = fill_valor_obligatorio if "Opcional" not in str(c_exp.value) else fill_valor
            c_valor.border = border_input
            c_valor.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            
            c_exp.font = font_explicacion
            c_exp.fill = fill_explicacion
            c_exp.border = border_standard
            c_exp.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

        dv_tipo = DataValidation(type="list", formula1='"F,J"', allow_blank=False)
        ws1.add_data_validation(dv_tipo)
        dv_tipo.add("B8")
        dv_tipo.add("B18")

        dv_iva = DataValidation(type="list", formula1='"21,10,4,0"', allow_blank=False)
        ws1.add_data_validation(dv_iva)
        dv_iva.add("B5")

        dv_irpf = DataValidation(type="list", formula1='"15,7,0"', allow_blank=False)
        ws1.add_data_validation(dv_irpf)
        dv_irpf.add("B6")

        ws2 = wb["Lineas_Factura"]
        ws2.views.sheetView[0].showGridLines = True
        ws2.freeze_panes = "A2"
        ws2.column_dimensions["A"].width = 58
        ws2.column_dimensions["B"].width = 16
        ws2.column_dimensions["C"].width = 22
        ws2.column_dimensions["D"].width = 24
        ws2.row_dimensions[1].height = 30

        headers_ws2 = ["Concepto", "Cantidad", "Precio_Unitario", "Subtotal (Automático)"]
        for col_idx, text in enumerate(headers_ws2, 1):
            cell = ws2.cell(row=1, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="left" if col_idx == 1 else "center", vertical="center")

        for row_idx in range(2, 17):
            ws2.row_dimensions[row_idx].height = 24
            c_con = ws2.cell(row=row_idx, column=1)
            c_can = ws2.cell(row=row_idx, column=2)
            c_pre = ws2.cell(row=row_idx, column=3)
            c_sub = ws2.cell(row=row_idx, column=4)
            
            c_con.font = Font(name="Calibri", size=11, color="0F172A")
            c_con.border = border_standard
            c_con.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            
            c_can.font = Font(name="Calibri", size=11, bold=True, color="0369A1")
            c_can.border = border_standard
            c_can.alignment = Alignment(horizontal="center", vertical="center")
            c_can.number_format = "#,##0.00"
            
            c_pre.font = Font(name="Calibri", size=11, bold=True, color="0369A1")
            c_pre.border = border_standard
            c_pre.alignment = Alignment(horizontal="right", vertical="center")
            c_pre.number_format = '#,##0.00 "€"'
            
            c_sub.value = f'=IF(OR(ISBLANK(B{row_idx}),ISBLANK(C{row_idx})),"",B{row_idx}*C{row_idx})'
            c_sub.font = Font(name="Calibri", size=11, bold=True, color="334155")
            c_sub.fill = fill_campo
            c_sub.border = border_standard
            c_sub.alignment = Alignment(horizontal="right", vertical="center")
            c_sub.number_format = '#,##0.00 "€"'

    return output.getvalue()


# ==============================================================================
# INTEGRACIÓN AUTOFIRMA LOCAL
# ==============================================================================

def es_entorno_local_con_autofirma():
    """
    Detecta si la aplicación se ejecuta localmente en modo desarrollo con AutoFirma instalado
    en macOS, Windows o Linux.
    En servidores de producción (MODO_DESARROLLO=False), se desactiva siempre la ejecución de
    subprocesos y escrituras en disco del servidor para garantizar el cumplimiento del RGPD.
    """
    if not MODO_DESARROLLO:
        return False
    sistema = platform.system()
    try:
        if sistema == "Darwin":
            return Path("/Applications/AutoFirma.app").exists()
        elif sistema == "Windows":
            if shutil.which("AutoFirma") or shutil.which("AutoFirmaCommandLine"):
                return True
            rutas_win = [
                r"C:\Program Files\AutoFirma\AutoFirma\AutoFirma.exe",
                r"C:\Program Files (x86)\AutoFirma\AutoFirma\AutoFirma.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\AutoFirma\AutoFirma.exe"),
            ]
            return any(os.path.exists(exe) for exe in rutas_win)
        elif sistema == "Linux":
            if shutil.which("autofirma"):
                return True
            rutas_linux = [
                "/usr/bin/autofirma",
                "/usr/local/bin/autofirma",
                "/opt/AutoFirma/AutoFirma",
                "/opt/autofirma/autofirma",
            ]
            return any(os.path.exists(exe) for exe in rutas_linux)
    except Exception:
        pass
    return False

def exportar_xml_local(nombre_xml, xml_bytes):
    try:
        carpeta = Path.home() / "Downloads" / "Facturas_Facturae"
        carpeta.mkdir(parents=True, exist_ok=True)
    except Exception:
        carpeta = BASE_DIR / "Facturas_Facturae"
        carpeta.mkdir(parents=True, exist_ok=True)
    ruta_xml = carpeta / nombre_xml
    with open(ruta_xml, "wb") as f:
        f.write(xml_bytes)
    return ruta_xml

def abrir_autofirma_gui(ruta_xml):
    sistema = platform.system()
    try:
        if sistema == "Darwin":
            app_path = Path("/Applications/AutoFirma.app")
            if app_path.exists():
                subprocess.Popen(["open", "-a", "AutoFirma", str(ruta_xml)])
                subprocess.Popen(["open", "-R", str(ruta_xml)])
                return True, f"AutoFirma se está abriendo en tu Mac. Archivo listo en: {ruta_xml}"
            else:
                subprocess.Popen(["open", "-R", str(ruta_xml)])
                return False, f"Se ha abierto la carpeta con el archivo {ruta_xml.name}. Si no tienes AutoFirma instalado, descárgalo de firmaelectronica.gob.es."
        elif sistema == "Windows":
            rutas_win = [
                r"C:\Program Files\AutoFirma\AutoFirma\AutoFirma.exe",
                r"C:\Program Files (x86)\AutoFirma\AutoFirma\AutoFirma.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\AutoFirma\AutoFirma.exe"),
            ]
            exe_encontrado = shutil.which("AutoFirma")
            if not exe_encontrado:
                for exe in rutas_win:
                    if os.path.exists(exe):
                        exe_encontrado = exe
                        break
            if exe_encontrado:
                subprocess.Popen([exe_encontrado])
                subprocess.Popen(["explorer", "/select,", str(ruta_xml)])
                return True, f"AutoFirma abierto. Arrastra el archivo resaltado: {ruta_xml}"
            subprocess.Popen(["explorer", "/select,", str(ruta_xml)])
            return False, "Se ha abierto la carpeta con tu XML. Ábrelo con AutoFirma."
        elif sistema == "Linux":
            exe_linux = shutil.which("autofirma")
            if not exe_linux:
                rutas_linux = [
                    "/usr/bin/autofirma",
                    "/usr/local/bin/autofirma",
                    "/opt/AutoFirma/AutoFirma",
                    "/opt/autofirma/autofirma",
                ]
                for p in rutas_linux:
                    if os.path.exists(p):
                        exe_linux = p
                        break
            if exe_linux:
                subprocess.Popen([exe_linux, str(ruta_xml)])
                if shutil.which("xdg-open"):
                    subprocess.Popen(["xdg-open", str(ruta_xml.parent)])
                return True, f"AutoFirma se está abriendo en tu equipo Linux. Archivo listo en: {ruta_xml}"
            elif shutil.which("xdg-open"):
                subprocess.Popen(["xdg-open", str(ruta_xml.parent)])
                return False, f"Se ha abierto la carpeta con el archivo {ruta_xml.name}. Si no tienes AutoFirma instalado, descárgalo de firmaelectronica.gob.es."
            else:
                return False, f"Archivo guardado en: {ruta_xml}. Ábrelo con AutoFirma."
        else:
            return False, f"Archivo guardado en: {ruta_xml}. Ábrelo desde AutoFirma."
    except Exception:
        return False, "Error al lanzar la aplicación AutoFirma en este equipo."

def firmar_cli_autofirma(ruta_xml):
    sistema = platform.system()
    ruta_xsig = ruta_xml.with_suffix(".xsig")
    if ruta_xsig.exists():
        ruta_xsig.unlink()

    cmd = None
    if sistema == "Darwin":
        bin_mac = Path("/Applications/AutoFirma.app/Contents/MacOS/AutoFirma")
        if bin_mac.exists():
            cmd = [str(bin_mac), "sign", "-i", str(ruta_xml), "-o", str(ruta_xsig), "-format", "facturae", "-certgui"]
    elif sistema == "Windows":
        rutas_win_cli = [
            r"C:\Program Files\AutoFirma\AutoFirma\AutoFirmaCommandLine.exe",
            r"C:\Program Files\AutoFirma\AutoFirma\AutoFirma.exe",
            r"C:\Program Files (x86)\AutoFirma\AutoFirma\AutoFirmaCommandLine.exe",
            r"C:\Program Files (x86)\AutoFirma\AutoFirma\AutoFirma.exe",
        ]
        exe_cli = shutil.which("AutoFirmaCommandLine") or shutil.which("AutoFirma")
        if not exe_cli:
            for exe in rutas_win_cli:
                if os.path.exists(exe):
                    exe_cli = exe
                    break
        if exe_cli:
            cmd = [exe_cli, "sign", "-i", str(ruta_xml), "-o", str(ruta_xsig), "-format", "facturae", "-certgui"]
    elif sistema == "Linux":
        exe_linux = shutil.which("autofirma")
        if not exe_linux:
            rutas_linux = [
                "/usr/bin/autofirma",
                "/usr/local/bin/autofirma",
                "/opt/AutoFirma/AutoFirma",
                "/opt/autofirma/autofirma",
            ]
            for p in rutas_linux:
                if os.path.exists(p):
                    exe_linux = p
                    break
        if exe_linux:
            cmd = [exe_linux, "sign", "-i", str(ruta_xml), "-o", str(ruta_xsig), "-format", "facturae", "-certgui"]

    if not cmd:
        return False, None, "No se detectó AutoFirma oficial en la ruta del sistema. Usa el botón de descarga del XML."

    try:
        subprocess.run(cmd, check=True, timeout=120)
        if ruta_xsig.exists():
            return True, ruta_xsig, f"¡Factura firmada (.xsig) creada con éxito en: {ruta_xsig}!"
        return False, None, "Se cerró la ventana de firma sin completar el proceso."
    except Exception:
        return False, None, "No se pudo completar la firma automática. Descarga el XML y fírmalo manualmente en AutoFirma."


# ==============================================================================
# ESTILOS E INTERFAZ GRÁFICA (CSS PULIDO Y AUDITADO)
# ==============================================================================

def inyectar_estilos():
    st.markdown(
        """
        <style>
        @import url("https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap");

        /* Tipografía cross-platform universal: macOS, iOS, Windows, Linux, Android */
        .stApp {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        /* Espaciado del contenedor principal: suficiente margen superior para que no colisione con el header ni con la flecha */
        .block-container {
            max-width: 1120px;
            padding-top: 3.6rem !important;
            padding-bottom: 3rem !important;
        }

        /* Header transparente sin tapar botones de colapso */
        header[data-testid="stHeader"] {
            background: transparent !important;
            z-index: 99 !important;
        }

        /* Botón de la flecha para colapsar/expandir el menú lateral: posición clara, visible y sin cortes */
        [data-testid="collapsedControl"] {
            top: 0.85rem !important;
            left: 0.85rem !important;
            z-index: 100 !important;
        }
        [data-testid="collapsedControl"] button,
        [data-testid="stSidebarCollapseButton"] button {
            border-radius: 8px !important;
            background: var(--secondary-background-color) !important;
            border: 1px solid rgba(128, 128, 128, 0.2) !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 14%, transparent) !important;
            box-shadow: 0 2px 6px rgba(0,0,0,0.06) !important;
            transition: all 0.2s ease !important;
        }

        footer { visibility: hidden; height: 0; }

        /* Estilo elegante para las pestañas con compatibilidad cross-browser */
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            gap: 0.35rem;
            background: var(--secondary-background-color);
            padding: 0.35rem;
            border-radius: 14px;
            border: 1px solid rgba(128, 128, 128, 0.18);
            border: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent);
        }
        div[data-testid="stTabs"] [data-baseweb="tab"] {
            border-radius: 10px;
            font-weight: 600;
            padding: 0.65rem 1.05rem;
            color: var(--text-color);
        }
        div[data-testid="stTabs"] [aria-selected="true"] {
            background: rgba(37, 99, 235, 0.12);
            background: color-mix(in srgb, #2563eb 16%, var(--background-color));
            color: var(--text-color);
        }

        /* Hero header moderno y responsive */
        .hero {
            display: flex;
            flex-wrap: wrap;
            gap: 1.25rem;
            align-items: stretch;
            justify-content: space-between;
            margin: 0.25rem 0 1.35rem 0;
        }
        .hero-copy { flex: 1 1 520px; }
        .kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            font-size: 0.72rem;
            font-weight: 700;
            color: #2563eb;
            margin: 0 0 0.55rem 0;
        }
        [data-theme="dark"] .kicker { color: #93c5fd; }
        .hero h1 {
            font-size: clamp(1.7rem, 3vw, 2.35rem);
            line-height: 1.18;
            font-weight: 800;
            letter-spacing: -0.03em;
            color: var(--text-color);
            margin: 0 0 0.65rem 0;
        }
        .lede {
            font-size: 1.02rem;
            line-height: 1.55;
            color: var(--text-color);
            opacity: 0.82;
            margin: 0;
            max-width: 46rem;
        }
        .compliance {
            flex: 0 1 280px;
            background: var(--secondary-background-color);
            border: 1px solid rgba(128, 128, 128, 0.2);
            border: 1px solid color-mix(in srgb, var(--text-color) 14%, transparent);
            border-radius: 16px;
            padding: 1.05rem 1.15rem;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.05);
            box-shadow: 0 8px 24px color-mix(in srgb, var(--text-color) 6%, transparent);
        }
        .compliance strong {
            display: block;
            color: var(--text-color);
            font-size: 0.92rem;
            margin-bottom: 0.45rem;
        }
        .compliance p {
            margin: 0;
            color: var(--text-color);
            opacity: 0.82;
            font-size: 0.88rem;
            line-height: 1.45;
        }
        .pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            border-radius: 999px;
            padding: 0.22rem 0.7rem;
            font-size: 0.75rem;
            font-weight: 700;
            margin-bottom: 0.7rem;
            background: rgba(34, 197, 94, 0.15);
            background: color-mix(in srgb, #22c55e 18%, transparent);
            color: var(--text-color);
            border: 1px solid rgba(34, 197, 94, 0.35);
            border: 1px solid color-mix(in srgb, #22c55e 42%, transparent);
        }
        .step-row {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.7rem;
            margin: 0 0 1.2rem 0;
        }
        .step {
            background: var(--secondary-background-color);
            border: 1px solid rgba(128, 128, 128, 0.18);
            border: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent);
            border-radius: 14px;
            padding: 0.85rem 0.9rem;
        }
        .step b { display: block; color: var(--text-color); font-size: 0.88rem; margin: 0.15rem 0 0.2rem 0; }
        .step span { color: var(--text-color); opacity: 0.75; font-size: 0.78rem; line-height: 1.35; }
        .step i { font-style: normal; font-size: 1.05rem; }
        @media (max-width: 820px) {
            .step-row { grid-template-columns: 1fr 1fr; }
        }
        @media (max-width: 540px) {
            .step-row { grid-template-columns: 1fr; }
            .hero { flex-direction: column; }
            .compliance { flex: 1 1 100%; }
        }
        .card-kicker {
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #2563eb;
            margin-bottom: 0.15rem;
        }
        [data-theme="dark"] .card-kicker { color: #93c5fd; }
        .legal-foot {
            margin-top: 1.4rem;
            padding-top: 0.9rem;
            border-top: 1px solid rgba(128, 128, 128, 0.18);
            border-top: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent);
            color: var(--text-color);
            opacity: 0.72;
            font-size: 0.78rem;
            line-height: 1.45;
        }
        div[data-testid="stMetric"] {
            background: var(--secondary-background-color);
            border: 1px solid rgba(128, 128, 128, 0.18);
            border: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent);
            border-radius: 14px;
            padding: 0.65rem 0.8rem;
        }
        div[data-testid="stMetricValue"] { color: var(--text-color); }
        div[data-testid="stMetricLabel"] { color: var(--text-color); opacity: 0.75; }

        /* Tarjeta PRO destacada */
        .pro-card {
            background: rgba(34, 197, 94, 0.12);
            background: color-mix(in srgb, #22c55e 14%, var(--secondary-background-color));
            border: 1px solid rgba(34, 197, 94, 0.35);
            border: 1px solid color-mix(in srgb, #22c55e 40%, transparent);
            border-radius: 14px;
            padding: 1rem 1.1rem;
            margin: 0.5rem 0 0.9rem 0;
        }
        .pro-card h4 { margin: 0 0 0.3rem 0; color: var(--text-color); font-weight: 700; font-size: 0.98rem; }
        .pro-card p { margin: 0; font-size: 0.82rem; opacity: 0.92; line-height: 1.4; }

        /* Banner de aviso de pago (paywall) */
        .paywall-banner {
            background: rgba(245, 158, 11, 0.12);
            background: color-mix(in srgb, #f59e0b 14%, var(--secondary-background-color));
            border: 1px solid rgba(245, 158, 11, 0.35);
            border: 1px solid color-mix(in srgb, #f59e0b 45%, transparent);
            border-radius: 16px;
            padding: 1.2rem 1.4rem;
            margin: 1rem 0;
        }
        .paywall-banner h4 { margin: 0 0 0.4rem 0; color: var(--text-color); font-weight: 700; }
        .paywall-banner p { margin: 0 0 0.8rem 0; font-size: 0.92rem; opacity: 0.9; }

        /* Estilo mejorado para el uploader (evita duplicidades visuales y bordes toscos) */
        [data-testid="stFileUploader"] {
            margin-top: 0.4rem;
        }
        [data-testid="stFileUploaderDropzone"] {
            padding: 1.2rem 1rem !important;
            border-radius: 12px !important;
            border: 2px dashed rgba(128, 128, 128, 0.28) !important;
            border: 2px dashed color-mix(in srgb, var(--text-color) 20%, transparent) !important;
            background: rgba(128, 128, 128, 0.05) !important;
            background: color-mix(in srgb, var(--secondary-background-color) 60%, transparent) !important;
            transition: border-color 0.2s ease;
        }
        [data-testid="stFileUploaderDropzone"]:hover {
            border-color: #2563eb !important;
        }

        /* Estilo limpio para los campos de texto y contraseña en la barra lateral */
        [data-testid="stSidebar"] [data-baseweb="input"] {
            border-radius: 8px !important;
        }
        [data-testid="stSidebar"] form {
            border: none !important;
            padding: 0 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_hero():
    st.markdown(
        """
        <div class="hero">
          <div class="hero-copy">
            <p class="kicker">Facturae B2B · Procesamiento 100% en Memoria</p>
            <h1>Convierte tu Excel en Facturae 3.2.2 listo para firmar</h1>
            <p class="lede">
              Plantilla guiada, validación fiscal y XML oficial para FACe y B2B.
              Procesamiento volátil en memoria RAM: ningún dato fiscal se almacena en el disco del servidor.
            </p>
          </div>
          <div class="compliance">
            <div class="pill">● Cumplimiento normativo</div>
            <strong>Ley Crea y Crece · FACe · Facturae 3.2.2</strong>
            <p>Esquema público XML 3.2.2. No es un Sistema Informático de Facturación (RD 1007/2023). Privacidad y cifrado de cuentas.</p>
          </div>
        </div>
        <div class="step-row">
          <div class="step"><i>①</i><b>Plantilla</b><span>Excel descargable con desplegables.</span></div>
          <div class="step"><i>②</i><b>Validar</b><span>Validación en memoria (NIF, IBAN, F/J).</span></div>
          <div class="step"><i>③</i><b>XML</b><span>Facturae 3.2.2 generado en RAM.</span></div>
          <div class="step"><i>④</i><b>Firmar</b><span>Descarga directa o AutoFirma local.</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# GESTIÓN DE USUARIOS, VERIFICACIÓN OTP Y CANJEO EN BARRA LATERAL
# ==============================================================================

def render_sidebar():
    with st.sidebar:
        user = st.session_state.usuario_actual
        pendiente = st.session_state.pendiente_verificacion
        
        # --- CASO A: PENDIENTE DE VERIFICAR CORREO (PANTALLA OTP) ---
        if pendiente:
            st.markdown("### 🔐 Verificar Correo")
            st.caption(f"Hemos enviado un código a: **{pendiente['email']}**")
            
            # Si no hay servidor SMTP, solo mostramos el código en pantalla si MODO_DESARROLLO está explícitamente activado
            if not pendiente.get("enviado_smtp"):
                if MODO_DESARROLLO:
                    st.info(f"🧪 **[Modo Pruebas / Sin SMTP]**:\nTu código de verificación es: **{pendiente.get('codigo')}**\n*(En producción este código se envía por correo vía TLS).*")
                else:
                    det = pendiente.get("detalle_smtp", "")
                    if det == "SMTP_NO_CONFIGURADO":
                        st.error("⚠️ El servidor de correo no está configurado. Por seguridad fiscal, en producción se requiere SMTP activo para verificar cuentas. Contacta con soporte.")
                    else:
                        st.error(f"⚠️ {det}")
            else:
                st.success("📧 Código enviado a tu buzón. Revisa también tu carpeta de spam.")

            with st.form("form_otp_verificar"):
                codigo_otp = st.text_input(
                    "Código de 6 dígitos",
                    placeholder="Ej: 492104",
                    max_chars=6,
                    help="Introduce el código numérico recibido por email."
                )
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    btn_confirmar_otp = st.form_submit_button("Verificar", use_container_width=True, type="primary")
                with col_b2:
                    btn_reenviar_otp = st.form_submit_button("Reenviar", use_container_width=True)

            if btn_confirmar_otp:
                if not codigo_otp.strip():
                    st.warning("Introduce los 6 dígitos del código.")
                else:
                    ok_otp, msg_otp = validar_codigo_otp(pendiente["email"], codigo_otp)
                    if ok_otp:
                        st.success(msg_otp)
                        st.session_state.usuario_actual = recargar_usuario(pendiente["email"])
                        st.session_state.pendiente_verificacion = None
                        st.rerun()
                    else:
                        st.error(msg_otp)

            if btn_reenviar_otp:
                cod, ok_m, det_m = generar_y_enviar_codigo_verificacion(pendiente["email"])
                st.session_state.pendiente_verificacion["codigo"] = cod
                st.session_state.pendiente_verificacion["enviado_smtp"] = ok_m
                st.session_state.pendiente_verificacion["detalle_smtp"] = det_m
                st.success("Nuevo código enviado.")
                st.rerun()

            if st.button("⬅️ Cambiar de Correo", use_container_width=True):
                st.session_state.pendiente_verificacion = None
                st.rerun()

        # --- CASO B: USUARIO AUTENTICADO Y VERIFICADO ---
        elif user:
            st.markdown("### 👤 Mi Cuenta")
            st.caption(f"Conectado como: **{user['email']}**")
            
            # --- USUARIO PRO ---
            if user["es_pro"]:
                tipo_lic = user.get("tipo_licencia") or "vitalicia"
                if tipo_lic == "vitalicia":
                    st.markdown(
                        """
                        <div class="pro-card">
                            <h4>👑 ¡Eres usuario PRO Vitalicio!</h4>
                            <p>Tu cuenta dispone de <b>conversiones y descargas ilimitadas de por vida</b> (Pago único). ¡Gracias por confiar en Facturae B2B!</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        """
                        <div class="pro-card">
                            <h4>💳 Suscripción PRO Mensual Activa</h4>
                            <p>Tu cuenta dispone de <b>conversiones y descargas ilimitadas</b> con facturación mensual (9,49 €/mes).</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    # Botón Stripe Customer Portal para gestionar / cancelar la suscripción mensual
                    if stripe and STRIPE_SECRET_KEY and stripe_billing:
                        portal_url, err_p = stripe_billing.crear_sesion_portal_cliente(
                            user["email"],
                            STRIPE_SECRET_KEY,
                            return_url=APP_BASE_URL,
                        )
                        if portal_url:
                            st.link_button("⚙️ Gestionar o Cancelar Suscripción", portal_url, use_container_width=True)
                            st.caption("🔒 Accede al portal oficial de Stripe para cambiar tu tarjeta o cancelar.")
            
            # --- USUARIO ESTÁNDAR (GRATUITO) ---
            else:
                restantes = max(0, 1 - user["facturas_usadas"])
                if restantes > 0:
                    st.info(f"🎁 **Prueba gratuita activa**: Te queda {restantes} factura de prueba.")
                else:
                    st.warning("⚠️ **Prueba gratuita consumida**")
                
                st.markdown("---")
                # Selector de planes Stripe (activación automática tras el pago)
                st.markdown("#### 💎 Adquirir Licencia PRO")
                st.caption("Activación automática al instante tras completar el pago:")
                render_planes_pago_stripe(user["email"], key_prefix="sb", stack=True)


            
            st.markdown("---")
            if st.button("🚪 Cerrar Sesión", use_container_width=True):
                st.session_state.usuario_actual = None
                st.rerun()

        # --- CASO C: USUARIO NO CONECTADO (INVITADO) ---
        else:
            st.markdown("### 🔐 Acceso a tu Cuenta")
            st.caption("Crea una cuenta verificada para proteger tus facturas y activar tu licencia PRO.")
            
            tab_login, tab_registro = st.tabs(["Iniciar Sesión", "Crear Cuenta"])
            
            with tab_login:
                with st.form("form_login"):
                    login_email = st.text_input("Correo electrónico", key="login_email")
                    login_pass = st.text_input("Contraseña", type="password", key="login_pass")
                    btn_login = st.form_submit_button("Entrar", use_container_width=True, type="primary")
                    
                    if btn_login:
                        if not login_email or not login_pass:
                            st.warning("Completa tu correo y contraseña.")
                        else:
                            ok_l, u_data, msg_l = autenticar_usuario(login_email, login_pass)
                            if ok_l:
                                st.session_state.usuario_actual = u_data
                                st.success(msg_l)
                                st.rerun()
                            elif isinstance(u_data, dict) and u_data.get("necesita_verificar"):
                                st.session_state.pendiente_verificacion = u_data
                                st.warning(msg_l)
                                st.rerun()
                            else:
                                st.error(msg_l)
            
            with tab_registro:
                with st.form("form_registro"):
                    reg_email = st.text_input("Tu Correo Real", key="reg_email", placeholder="nombre@empresa.com")
                    reg_pass = st.text_input(
                        "Contraseña",
                        type="password",
                        key="reg_pass",
                        help="Mínimo 8 caracteres combinando letras y números."
                    )
                    reg_pass2 = st.text_input("Repite la contraseña", type="password", key="reg_pass2")
                    btn_reg = st.form_submit_button("Crear Cuenta y Verificar", use_container_width=True, type="primary")
                    
                    if btn_reg:
                        if reg_pass != reg_pass2:
                            st.error("Las contraseñas no coinciden.")
                        else:
                            ok_r, res_r = registrar_usuario(reg_email, reg_pass)
                            if ok_r:
                                st.session_state.pendiente_verificacion = res_r
                                st.success("Cuenta registrada. Introduce el código enviado a tu correo.")
                                st.rerun()
                            else:
                                st.error(res_r)

            st.markdown("---")
            st.markdown("#### 💎 Adquirir Licencia PRO")
            st.caption("Elige tu modalidad (activación inmediata):")
            render_planes_pago_stripe(email="", key_prefix="guest_sb", stack=True)



# ==============================================================================
# FLUJO PRINCIPAL DE LA APLICACIÓN
# ==============================================================================

inyectar_estilos()
render_sidebar()
render_hero()

# ==============================================================================
# PROCESAMIENTO AUTOMÁTICO DE RETORNO DE STRIPE (?session_id=...)
# ==============================================================================
query_session_id = st.query_params.get("session_id")
if query_session_id:
    user_actual = st.session_state.usuario_actual
    if user_actual:
        if user_actual.get("es_pro"):
            st.query_params.clear()
        else:
            ok_stripe, msg_stripe = verificar_y_activar_sesion_stripe(query_session_id, user_actual["email"])
            st.query_params.clear()
            if ok_stripe:
                st.session_state.usuario_actual = recargar_usuario(user_actual["email"])
                st.balloons()
                st.success(msg_stripe)
                st.rerun()
            else:
                st.error(msg_stripe)
    else:
        if stripe and STRIPE_SECRET_KEY:
            try:
                stripe.api_key = STRIPE_SECRET_KEY
                sess_info = stripe.checkout.Session.retrieve(query_session_id)
                email_comprador = (
                    sess_info.client_reference_id
                    or (sess_info.customer_details and sess_info.customer_details.email)
                    or ""
                ).strip().lower()
                if email_comprador:
                    user_existente = recargar_usuario(email_comprador)
                    if user_existente:
                        ok_s, msg_s = verificar_y_activar_sesion_stripe(query_session_id, email_comprador)
                        if ok_s:
                            st.session_state.usuario_actual = recargar_usuario(email_comprador)
                            st.query_params.clear()
                            st.balloons()
                            st.success(f"{msg_s} ¡Sesión iniciada automáticamente como **{email_comprador}**!")
                            st.rerun()
                    else:
                        st.success(f"🎉 ¡Pago completado con éxito para **{email_comprador}**!")
                        st.info("Crea tu cuenta gratuita en la barra lateral con este mismo correo para activar tu cuenta PRO al instante.")
            except Exception:
                pass
        st.warning("⚠️ Detectamos un retorno de pago de Stripe. Inicia sesión en el menú lateral con la cuenta con la que compraste para completar la activación de tu licencia PRO.")


tab_convertir, tab_planes, tab_guia, tab_soporte = st.tabs(
    ["⚡ Convertir y Firmar", "💎 Planes y Precios", "📘 Guía y Requisitos", "🛟 Soporte y Errores"]
)


with tab_convertir:
    with st.container(border=True):
        st.markdown('<p class="card-kicker">Antes de continuar</p>', unsafe_allow_html=True)
        st.markdown("##### ⚖️ Exención de responsabilidad legal")
        st.caption(
            "Esta herramienta realiza exclusivamente la conversión sintáctica de los datos de la plantilla Excel al formato "
            "XML Facturae 3.2.2 en memoria volátil. No es un programa de contabilidad ni un Sistema Informático de Facturación "
            "(RD 1007/2023), ni conserva datos en ningún servidor. El usuario es el único responsable de comprobar la exactitud "
            "de los importes, NIF y datos fiscales antes de presentarlos a Hacienda o FACe."
        )
        acepto_aviso = st.checkbox(
            "He leído el aviso y entiendo que soy responsable de la veracidad de los datos de mi factura.",
            value=True,
        )

    col_plantilla, col_subida = st.columns(2, gap="large")
    with col_plantilla:
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Paso 1</p>', unsafe_allow_html=True)
            st.markdown("##### 📥 Descargar Plantilla")
            st.write(
                "Plantilla oficial guiada con desplegables fiscales (IVA, IRPF, F/J) y fórmulas automáticas. "
                "Descárgala y complétala en tu ordenador."
            )
            st.download_button(
                label="Descargar Plantilla_Facturae.xlsx",
                data=generar_plantilla_excel(),
                file_name="Plantilla_Facturae.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

    with col_subida:
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Paso 2</p>', unsafe_allow_html=True)
            st.markdown("##### 📤 Subir Archivo Rellenado")
            st.caption("Selecciona o arrastra la plantilla `.xlsx` completada (procesamiento 100% en RAM):")
            archivo = st.file_uploader(
                "Selecciona tu plantilla Excel",
                type=["xlsx"],
                label_visibility="collapsed",
                help="Arrastra aquí el archivo Excel completado."
            )

    factura_lista = False
    xml_bytes = b""
    nombre_xml = "Factura.xml"

    with st.container(border=True):
        st.markdown('<p class="card-kicker">Paso 3</p>', unsafe_allow_html=True)
        st.markdown("##### 🧮 Validación y desglose fiscal")
        if not archivo:
            st.info("Sube tu archivo Excel para validar los datos fiscales y calcular la base, IVA, IRPF y total.")
        else:
            try:
                contenido_bytes = archivo.getvalue()
                
                # Validación de seguridad: tamaño máximo de archivo
                if len(contenido_bytes) > MAX_FILE_SIZE_BYTES:
                    st.error("El archivo excede el tamaño máximo de seguridad permitido (5 MB).")
                else:
                    buffer_excel = BytesIO(contenido_bytes)
                    df_cab = pd.read_excel(buffer_excel, sheet_name="Datos_Factura")
                    buffer_excel.seek(0)
                    df_lin = pd.read_excel(buffer_excel, sheet_name="Lineas_Factura")
                    del buffer_excel  # Destrucción en RAM

                    errores, datos_dict = validar_datos(df_cab, df_lin)
                    if errores:
                        st.error("Se han detectado campos pendientes o formatos erróneos en el Excel:")
                        for err in errores:
                            st.write(f"• {err}")
                    else:
                        xml_bytes, base, iva, irpf, total, df_preview = construir_xml_facturae(datos_dict, df_lin)
                        nombre_xml = f"Factura_{limpiar_texto(datos_dict.get('Serie_Factura'))}_{limpiar_texto(datos_dict.get('Numero_Factura'))}.xml"
                        factura_lista = True

                        st.success("✅ Datos validados con éxito. Resumen contable de la factura:")
                        col1, col2, col3, col4 = st.columns(4)
                        col1.metric("Base imponible", f"{base:.2f} €")
                        col2.metric("IVA (+)", f"{iva:.2f} €")
                        col3.metric("IRPF (−)", f"{irpf:.2f} €")
                        col4.metric("Total factura", f"{total:.2f} €")

                        with st.expander("Ver desglose de conceptos facturados"):
                            st.dataframe(df_preview, use_container_width=True, hide_index=True)
            except Exception:
                st.error("El archivo no es válido o está corrupto. Asegúrate de subir la plantilla oficial sin macros.")

    with st.container(border=True):
        st.markdown('<p class="card-kicker">Paso 4</p>', unsafe_allow_html=True)
        st.markdown("##### 🔏 Descargar XML y firmar con AutoFirma")
        st.write(
            "Con el Excel validado y el aviso aceptado, descarga el `.xml` oficial generado en memoria o ábrelo en AutoFirma."
        )

        if not acepto_aviso:
            st.warning("Marca la casilla de exención de responsabilidad para habilitar la descarga.")

        # Lógica de autorización y control de licencias (Blindaje Anti-Bypass)
        usuario = st.session_state.usuario_actual
        es_pro = bool(usuario and usuario.get("es_pro"))
        esta_autenticado = usuario is not None
        
        if not esta_autenticado:
            st.markdown(
                """
                <div class="paywall-banner" style="border-left: 4px solid #2563eb;">
                  <h4>🔐 Identificación requerida para descargar tu factura</h4>
                  <p>Por confidencialidad fiscal y para activar tu <b>1ª prueba gratuita</b> (o acceder a tu cuenta PRO ilimitada), <b>inicia sesión o crea tu cuenta gratuita</b> en el menú lateral. El registro toma menos de 30 segundos.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            render_planes_pago_stripe(email="", key_prefix="paso4_guest", stack=False)
            bloqueado_por_licencia = True

        else:
            veces_generadas = usuario.get("facturas_usadas", 0)
            bloqueado_por_licencia = (not es_pro) and (veces_generadas >= 1)

            # Si el usuario es PRO, mostramos confirmación
            if es_pro:
                st.markdown(
                    """
                    <div class="pro-card">
                        <h4>👑 Cuenta PRO Activa</h4>
                        <p>Tienes acceso ilimitado para descargar y firmar todas tus facturas electrónicas.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            elif bloqueado_por_licencia:
                st.markdown(
                    """
                    <div class="paywall-banner">
                      <h4>🔒 Límite de prueba gratuita alcanzado</h4>
                      <p>Has consumido tu <b>1 prueba gratuita</b>. Consigue acceso ilimitado eligiendo la <b>suscripción mensual de 9,49 €/mes</b> (cancela cuando quieras) o la <b>licencia vitalicia de 99,49 €</b> (pago único de por vida).</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                render_planes_pago_stripe(usuario["email"], key_prefix="paso4")
            else:
                st.info("🎁 **Prueba gratuita activa**: Tienes 1 descarga de factura incluida sin coste.")

        botones_deshabilitados = (not factura_lista) or (not acepto_aviso) or bloqueado_por_licencia

        # Descarga directa 100% desde la memoria RAM (Volátil y Segura)
        # BLINDAJE CRÍTICO RED TEAM: data es b"" si el botón está deshabilitado
        # para impedir la extracción del payload vía DevTools / DOM / WebSocket.
        payload_descarga = xml_bytes if (factura_lista and not botones_deshabilitados) else b""

        st.download_button(
            label=f"⬇️ Descargar {nombre_xml}",
            data=payload_descarga,
            file_name=nombre_xml,
            mime="application/xml",
            use_container_width=True,
            type="primary",
            disabled=botones_deshabilitados,
            on_click=registrar_factura_generada,
        )

        # Firma con AutoFirma: Detección inteligente entre escritorio local vs servidor cloud
        if es_entorno_local_con_autofirma():
            col_af1, col_af2 = st.columns(2)
            with col_af1:
                if st.button(
                    "🔐 Enviar y abrir en AutoFirma",
                    use_container_width=True,
                    disabled=botones_deshabilitados,
                ):
                    registrar_factura_generada()
                    ruta_local = exportar_xml_local(nombre_xml, xml_bytes)
                    ok, msg = abrir_autofirma_gui(ruta_local)
                    if ok:
                        st.success(msg)
                    else:
                        st.info(msg)

            with col_af2:
                if st.button(
                    "⚡ Firmar en 1 clic (.xsig)",
                    use_container_width=True,
                    disabled=botones_deshabilitados,
                ):
                    registrar_factura_generada()
                    ruta_local = exportar_xml_local(nombre_xml, xml_bytes)
                    with st.spinner("Abriendo diálogo de selección de certificado..."):
                        ok_cli, ruta_xsig, msg_cli = firmar_cli_autofirma(ruta_local)
                    if ok_cli and ruta_xsig:
                        st.success(msg_cli)
                        with open(ruta_xsig, "rb") as f_xsig:
                            st.download_button(
                                label=f"✅ Descargar {ruta_xsig.name}",
                                data=f_xsig.read(),
                                file_name=ruta_xsig.name,
                                mime="application/octet-stream",
                                use_container_width=True,
                            )
                    else:
                        st.warning(msg_cli)
        else:
            st.markdown(
                """
                <div style="background: var(--secondary-background-color); border: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent); border-radius: 14px; padding: 1.1rem 1.3rem; margin-top: 1rem;">
                    <strong style="color: var(--text-color); font-size: 0.95rem; display: block; margin-bottom: 0.4rem;">🔏 Cómo firmar este archivo XML en AutoFirma:</strong>
                    <ol style="margin: 0; padding-left: 1.25rem; font-size: 0.88rem; line-height: 1.55; opacity: 0.9; color: var(--text-color);">
                        <li>Pulsa el botón superior <b>⬇️ Descargar XML</b> para guardar la factura validada en tu equipo.</li>
                        <li>Abre tu aplicación oficial <b>AutoFirma</b> (descargable gratis en firmaelectronica.gob.es).</li>
                        <li>Selecciona el archivo <code>.xml</code> descargado y pulsa <b>Firmar</b> eligiendo tu certificado digital (FNMT, DNIe o cameral).</li>
                        <li>Guarda el archivo firmado resultante con extensión <code>.xsig</code> listo para subir a FACe o enviar a tu cliente.</li>
                    </ol>
                </div>
                """,
                unsafe_allow_html=True,
            )

with tab_planes:
    st.markdown('<p class="card-kicker">Planes Oficiales Facturae B2B</p>', unsafe_allow_html=True)
    st.markdown("### 💎 Elige el plan que mejor se adapta a tu negocio")
    st.write(
        "Convierte ilimitadamente tus plantillas Excel a XML Facturae 3.2.2 oficial cumpliendo con "
        "la Ley Crea y Crece y FACe. Dispones de dos modalidades oficiales procesadas con total seguridad por Stripe:"
    )
    email_u = (st.session_state.usuario_actual.get("email") if st.session_state.usuario_actual else "")
    render_planes_pago_stripe(email_u, key_prefix="main_tab", stack=False)

with tab_guia:
    g1, g2 = st.columns(2, gap="large")
    with g1:
        with st.container(border=True):
            st.markdown('<p class="card-kicker">El formato</p>', unsafe_allow_html=True)
            st.markdown("##### 📄 Qué es Facturae 3.2.2")
            st.write(
                "Facturae es el estándar XML obligatorio para facturar a la Administración Pública (FACe) y relaciones B2B en España. "
                "La versión **3.2.2** estructura emisor, receptor, conceptos, impuestos e IBAN de cobro."
            )
            st.write(
                "La aplicación traduce en memoria volátil los datos del Excel a la estructura oficial para que la firmes con tu certificado."
            )
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Normativa</p>', unsafe_allow_html=True)
            st.markdown("##### 🏛️ Ley Crea y Crece y FACe")
            st.write(
                "La Ley Crea y Crece impulsa la factura electrónica en el tejido empresarial español. Si tu cliente es una entidad pública, "
                "debes remitirla firmada a través de **FACe** indicando los tres códigos **DIR3** obligatorios."
            )
            st.caption("Los campos DIR3 de la plantilla son opcionales: rellénalos solo si facturas a un organismo público.")
    with g2:
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Instrucciones</p>', unsafe_allow_html=True)
            st.markdown("##### 🧭 Cómo usar la herramienta")
            st.markdown(
                """
1. Descarga la plantilla y selecciona **F** (autónomo) o **J** (empresa).
2. Rellena los datos fiscales, domicilio, IBAN español y al menos un concepto en la hoja de líneas.
3. Sube el archivo Excel: el sistema valida campos e importes en memoria.
4. Acepta el aviso legal y descarga tu XML oficial o fírmalo con AutoFirma.
                """
            )
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Firma digital</p>', unsafe_allow_html=True)
            st.markdown("##### 🔐 AutoFirma")
            st.write(
                "AutoFirma es la aplicación oficial para firmar documentos digitales con certificados instalados en tu navegador o sistema. "
                "Puedes descargar el XML y firmarlo manualmente o utilizar los accesos locales directos."
            )

    with st.container(border=True):
        st.markdown('<p class="card-kicker">Requisitos</p>', unsafe_allow_html=True)
        st.markdown("##### ✅ Validación obligatoria")
        c_a, c_b, c_c = st.columns(3)
        with c_a:
            st.markdown("**Identidad F / J**")
            st.caption("F exige nombre y primer apellido. J exige razón social.")
        with c_b:
            st.markdown("**IBAN y fechas**")
            st.caption("IBAN español válido sin espacios. Fecha de emisión en formato AAAA-MM-DD.")
        with c_c:
            st.markdown("**Líneas de factura**")
            st.caption("Al menos un concepto con cantidad y precio. IVA (21/10/4/0) e IRPF (15/7/0).")

with tab_soporte:
    s1, s2 = st.columns((1.15, 0.85), gap="large")
    with s1:
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Centro de soporte</p>', unsafe_allow_html=True)
            st.markdown("##### 🙋‍♂️ Contactar con Soporte Técnico")
            st.caption(
                "¿Tienes alguna duda o incidencia con tu factura? Cuéntanos qué ha ocurrido "
                "y nuestro equipo revisará tu caso para ayudarte."
            )
            with st.form("form_reporte_error", clear_on_submit=True):
                categoria_reporte = st.selectbox("Tipo de consulta o problema", CATEGORIAS_REPORTE)
                descripcion_reporte = st.text_area(
                    "Mensaje o detalle de la consulta",
                    placeholder="Explica detalladamente qué necesitas o qué ha ocurrido...",
                    height=140,
                )
                enviado_reporte = st.form_submit_button("Enviar consulta", use_container_width=True, type="primary")

            if enviado_reporte:
                if categoria_reporte == CATEGORIAS_REPORTE[0]:
                    st.warning("Selecciona una categoría de problema antes de enviar.")
                elif not descripcion_reporte.strip():
                    st.warning("Introduce una descripción detallada del problema.")
                else:
                    ok_rep, msg_rep = guardar_reporte_error(categoria_reporte, descripcion_reporte)
                    if ok_rep:
                        st.success(msg_rep)
                    else:
                        st.error(msg_rep)
    with s2:
        with st.container(border=True):
            st.markdown('<p class="card-kicker">Soluciones rápidas</p>', unsafe_allow_html=True)
            st.markdown("##### 💡 Preguntas frecuentes")
            st.write("**Verificación de correo** — Revisa tu carpeta de spam o correo no deseado si no ves el código en pocos segundos.")
            st.write("**Error al subir Excel** — Asegúrate de mantener los nombres de las pestañas `Datos_Factura` y `Lineas_Factura`.")
            st.write("**Cálculo de impuestos** — Los subtotales se calculan automáticamente: Base + IVA − IRPF.")
            st.write("**Cuenta PRO** — Si ya compraste tu licencia, canjea tu código en el menú lateral para activar tu acceso ilimitado.")
            st.write("**Código ya usado** — Los códigos de licencia son de un único uso. Si tienes algún problema con tu compra, escríbenos desde este formulario.")

st.markdown(
    '<p class="legal-foot">Aviso de seguridad y privacidad: Esta utilidad realiza exclusivamente la traducción sintáctica en memoria RAM '
    "de archivos Excel al estándar oficial Facturae 3.2.2. No constituye un Sistema Informático de Facturación (RD 1007/2023) "
    "ni conserva datos fiscales en ningún servidor.</p>",
    unsafe_allow_html=True,
)
