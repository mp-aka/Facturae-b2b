"""
Módulo de Integración de Stripe Payments & Billing para FacturaeB2B.
Soporte para suscripciones recurrentes, gestión de clientes, Customer Portal y webhooks.
"""

import os
import sqlite3
from datetime import datetime
from typing import Optional, Tuple, Dict, Any
try:
    import stripe
except ImportError:
    stripe = None

DB_PATH_DEFAULT = os.path.join(os.path.dirname(__file__), "facturae_b2b.db")


def get_db_connection(db_path: str = DB_PATH_DEFAULT) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn


# Precios oficiales de Stripe
STRIPE_PRICE_VITALICIA = "price_1USGYQROobv6Wd8mcdXDWFl1"  # 99,49 € (Pago único)
STRIPE_PRICE_MENSUAL = "price_1USGYQROobv6Wd8mr2jwehcN"    # 9,49 € / mes (Suscripción)


def inicializar_schema_stripe(db_path: str = DB_PATH_DEFAULT):
    """
    Aplica migraciones seguras e idempotentes en SQLite para almacenar
    los identificadores de suscripción de Stripe y registro de eventos (anti-replay).
    """
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("PRAGMA synchronous=NORMAL;")

        # Comprobar columnas existentes en usuarios
        cursor.execute("PRAGMA table_info(usuarios)")
        columnas = {col[1] for col in cursor.fetchall()}

        if "stripe_customer_id" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN stripe_customer_id TEXT")
        if "stripe_subscription_id" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN stripe_subscription_id TEXT")
        if "subscription_status" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN subscription_status TEXT DEFAULT 'inactive'")
        if "subscription_period_end" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN subscription_period_end TEXT")
        if "tipo_licencia" not in columnas:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN tipo_licencia TEXT")

        # Tabla para deduplicación idempotente de eventos de webhooks
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stripe_eventos_procesados (
                event_id TEXT PRIMARY KEY,
                event_type TEXT NOT NULL,
                processed_at TEXT NOT NULL,
                payload_summary TEXT
            )
        """)

        # Tabla histórica de transacciones y pagos
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS transacciones_stripe (
                session_id TEXT PRIMARY KEY,
                email TEXT NOT NULL,
                monto_cents INTEGER,
                moneda TEXT,
                fecha_pago TEXT NOT NULL,
                customer_id TEXT,
                subscription_id TEXT
            )
        """)
        cursor.execute("PRAGMA table_info(transacciones_stripe)")
        tx_cols = {col[1] for col in cursor.fetchall()}
        if "customer_id" not in tx_cols:
            cursor.execute("ALTER TABLE transacciones_stripe ADD COLUMN customer_id TEXT")
        if "subscription_id" not in tx_cols:
            cursor.execute("ALTER TABLE transacciones_stripe ADD COLUMN subscription_id TEXT")
        conn.commit()


def obtener_o_crear_cliente(
    email: str,
    secret_key: str,
    nombre: Optional[str] = None,
    db_path: str = DB_PATH_DEFAULT,
) -> Tuple[Optional[str], Optional[str]]:
    """
    Recupera el stripe_customer_id de la base de datos local o de Stripe.
    Si no existe, crea un nuevo Customer en Stripe con metadatos asociados.
    """
    stripe.api_key = secret_key
    email_clean = email.strip().lower()

    # 1. Buscar en BD local
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT stripe_customer_id FROM usuarios WHERE email = ?", (email_clean,))
        row = cursor.fetchone()
        if row and row["stripe_customer_id"]:
            return row["stripe_customer_id"], None

    # 2. Buscar o crear en la API de Stripe
    try:
        customers = stripe.Customer.list(email=email_clean, limit=1)
        if customers.data:
            customer = customers.data[0]
        else:
            customer = stripe.Customer.create(
                email=email_clean,
                name=nombre or email_clean.split("@")[0],
                metadata={"app": "FacturaeB2B", "email": email_clean},
            )

        customer_id = customer.id

        # Persistir en BD local
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE usuarios SET stripe_customer_id = ? WHERE email = ?",
                (customer_id, email_clean),
            )
            conn.commit()

        return customer_id, None
    except Exception as e:
        return None, f"Error al gestionar el cliente en Stripe: {str(e)}"


def crear_checkout_suscripcion(
    email: str,
    secret_key: str,
    base_url: str = "http://localhost:8501",
    price_id: Optional[str] = None,
    unit_amount_cents: int = 900,  # 9,00 € mensual por defecto
    currency: str = "eur",
    product_name: str = "Facturae B2B · Suscripción PRO",
    db_path: str = DB_PATH_DEFAULT,
) -> Tuple[Optional[str], Optional[str]]:
    """
    Crea una sesión de Stripe Checkout en modo suscripción recurrente (Billing).
    Activa recolección de NIF/CIF (tax_id_collection) para empresas B2B.
    """
    if not secret_key:
        return None, "La clave secreta de Stripe no está configurada."

    stripe.api_key = secret_key
    customer_id, err = obtener_o_crear_cliente(email, secret_key, db_path=db_path)
    if err or not customer_id:
        return None, err or "No se pudo obtener el ID de cliente en Stripe."

    clean_base_url = (base_url or "http://localhost:8501").rstrip("/")
    success_url = f"{clean_base_url}/?session_id={{CHECKOUT_SESSION_ID}}&status=success"
    cancel_url = f"{clean_base_url}/?status=cancelled"

    try:
        # Configurar ítems: usar un Price ID existente o crear precio recurrente dinámico
        if price_id:
            line_items = [{"price": price_id, "quantity": 1}]
        else:
            line_items = [{
                "price_data": {
                    "currency": currency.lower(),
                    "unit_amount": unit_amount_cents,
                    "recurring": {"interval": "month"},
                    "product_data": {
                        "name": product_name,
                        "description": "Conversiones y firmas ilimitadas de facturas electrónicas Facturae B2B.",
                    },
                },
                "quantity": 1,
            }]

        session = stripe.checkout.Session.create(
            customer=customer_id,
            payment_method_types=["card"],
            line_items=line_items,
            mode="subscription",
            client_reference_id=email.strip().lower(),
            metadata={
                "email": email.strip().lower(),
                "producto": "FacturaeB2B_Suscripcion_PRO",
            },
            subscription_data={
                "metadata": {"email": email.strip().lower()}
            },
            # Características fiscales para B2B en España/UE
            tax_id_collection={"enabled": True},
            billing_address_collection="required",
            allow_promotion_codes=True,
            success_url=success_url,
            cancel_url=cancel_url,
        )
        return session.url, None
    except Exception as e:
        return None, f"No se pudo crear la sesión de Checkout: {str(e)}"


def crear_sesion_portal_cliente(
    email: str,
    secret_key: str,
    return_url: str = "http://localhost:8501",
    db_path: str = DB_PATH_DEFAULT,
) -> Tuple[Optional[str], Optional[str]]:
    """
    Crea una sesión del Portal de Clientes de Stripe (Stripe Billing Customer Portal)
    para que el usuario gestione su método de pago, descargue facturas o cancele.
    """
    if not secret_key:
        return None, "La clave secreta de Stripe no está configurada."

    stripe.api_key = secret_key
    customer_id, err = obtener_o_crear_cliente(email, secret_key, db_path=db_path)
    if err or not customer_id:
        return None, err or "No se encontró el identificador del cliente en Stripe."

    try:
        portal_session = stripe.billing_portal.Session.create(
            customer=customer_id,
            return_url=return_url,
        )
        return portal_session.url, None
    except Exception as e:
        return None, f"No se pudo abrir el portal de facturación: {str(e)}"


def procesar_evento_stripe_idempotente(
    event: Dict[str, Any],
    db_path: str = DB_PATH_DEFAULT,
) -> Tuple[bool, str]:
    """
    Procesador idempotente para eventos de Stripe Billing.
    Garantiza que un evento repetido (reintentos de Stripe) no provoque efectos duplicados.
    """
    event_id = event.get("id")
    event_type = event.get("type")
    data_object = event.get("data", {}).get("object", {})

    if not event_id or not event_type:
        return False, "Evento inválido sin ID o tipo."

    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()

        # 1. Verificación de Idempotencia
        cursor.execute("SELECT event_id FROM stripe_eventos_procesados WHERE event_id = ?", (event_id,))
        if cursor.fetchone():
            return True, f"Evento {event_id} ya fue procesado anteriormente (idempotencia confirmada)."

        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 2. Enrutamiento de eventos de Billing & Subscriptions
        if event_type == "checkout.session.completed":
            email = (
                data_object.get("client_reference_id")
                or data_object.get("customer_email")
                or data_object.get("customer_details", {}).get("email")
                or ""
            ).strip().lower()
            customer_id = data_object.get("customer")
            subscription_id = data_object.get("subscription")
            amount_total = data_object.get("amount_total")
            currency = data_object.get("currency")
            session_id = data_object.get("id")

            mode = data_object.get("mode")
            is_sub = (mode == "subscription") or bool(subscription_id)
            tipo_licencia = "mensual" if is_sub else "vitalicia"

            if email:
                cursor.execute("""
                    UPDATE usuarios
                    SET es_pro = 1,
                        tipo_licencia = ?,
                        fecha_pro = COALESCE(fecha_pro, ?),
                        stripe_customer_id = COALESCE(?, stripe_customer_id),
                        stripe_subscription_id = COALESCE(?, stripe_subscription_id),
                        subscription_status = 'active'
                    WHERE email = ?
                """, (tipo_licencia, ahora, customer_id, subscription_id, email))

                if session_id:
                    cursor.execute("""
                        INSERT OR REPLACE INTO transacciones_stripe
                        (session_id, email, monto_cents, moneda, fecha_pago, customer_id, subscription_id)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (session_id, email, amount_total, currency, ahora, customer_id, subscription_id))

        elif event_type == "invoice.payment_succeeded":
            customer_id = data_object.get("customer")
            subscription_id = data_object.get("subscription")
            lines = data_object.get("lines", {}).get("data", [])
            period_end = None
            if lines:
                period_end_ts = lines[0].get("period", {}).get("end")
                if period_end_ts:
                    period_end = datetime.fromtimestamp(period_end_ts).strftime("%Y-%m-%d %H:%M:%S")

            cursor.execute("""
                UPDATE usuarios
                SET es_pro = 1,
                    subscription_status = 'active',
                    subscription_period_end = COALESCE(?, subscription_period_end)
                WHERE stripe_customer_id = ? OR stripe_subscription_id = ?
            """, (period_end, customer_id, subscription_id))

        elif event_type == "invoice.payment_failed":
            # Si el pago mensual falla, se revoca el acceso PRO (si no es licencia vitalicia)
            customer_id = data_object.get("customer")
            subscription_id = data_object.get("subscription")
            cursor.execute("""
                UPDATE usuarios
                SET es_pro = 0,
                    subscription_status = 'past_due'
                WHERE (stripe_customer_id = ? OR stripe_subscription_id = ?)
                  AND (tipo_licencia IS NULL OR tipo_licencia != 'vitalicia')
            """, (customer_id, subscription_id))

        elif event_type == "customer.subscription.updated":
            subscription_id = data_object.get("id")
            customer_id = data_object.get("customer")
            status = data_object.get("status")  # active, past_due, canceled, incomplete
            current_period_end_ts = data_object.get("current_period_end")
            period_end = None
            if current_period_end_ts:
                period_end = datetime.fromtimestamp(current_period_end_ts).strftime("%Y-%m-%d %H:%M:%S")

            es_pro = 1 if status in ("active", "trialing") else 0
            cursor.execute("""
                UPDATE usuarios
                SET es_pro = ?,
                    subscription_status = ?,
                    subscription_period_end = COALESCE(?, subscription_period_end)
                WHERE (stripe_subscription_id = ? OR stripe_customer_id = ?)
                  AND (tipo_licencia IS NULL OR tipo_licencia != 'vitalicia')
            """, (es_pro, status, period_end, subscription_id, customer_id))

        elif event_type == "customer.subscription.deleted":
            # Si el usuario cancela la suscripción mensual, se revoca el PRO de inmediato
            subscription_id = data_object.get("id")
            customer_id = data_object.get("customer")
            cursor.execute("""
                UPDATE usuarios
                SET es_pro = 0,
                    subscription_status = 'canceled'
                WHERE (stripe_subscription_id = ? OR stripe_customer_id = ?)
                  AND (tipo_licencia IS NULL OR tipo_licencia != 'vitalicia')
            """, (subscription_id, customer_id))

        # 3. Registrar evento en tabla de idempotencia
        cursor.execute("""
            INSERT INTO stripe_eventos_procesados (event_id, event_type, processed_at, payload_summary)
            VALUES (?, ?, ?, ?)
        """, (event_id, event_type, ahora, f"status={event_type}"))

        conn.commit()

    return True, f"Evento {event_type} ({event_id}) procesado correctamente."
