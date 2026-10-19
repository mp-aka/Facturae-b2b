"""
Managed Payments Integration Module.
Implements the Stripe Managed Payments blueprint:
- Creates recurring digital subscription product with tax_code txcd_10103100 and API version 2026-02-25.preview
- Generates Checkout Session with managed_payments[enabled]=true and mode=subscription
- Handles checkout.session.completed webhook events and persists domain data
"""

import os
import sqlite3
import logging
from datetime import datetime
from typing import Optional, Tuple, Dict, Any
try:
    import stripe
except ImportError:
    stripe = None

logger = logging.getLogger("ManagedPayments")

# Blueprint Constants
STRIPE_API_VERSION_PREVIEW = "2026-02-25.preview"
ONE_TIME_PRODUCT_NAME = "Hamlet (e-book)"
ONE_TIME_TAX_CODE = "txcd_10103100"
ONE_TIME_CURRENCY = "usd"
ONE_TIME_UNIT_AMOUNT = 1000  # $10.00

SUBSCRIPTION_PRODUCT_NAME = "Basic subscription"
SUBSCRIPTION_TAX_CODE = "txcd_10103100"
SUBSCRIPTION_CURRENCY = "usd"
SUBSCRIPTION_UNIT_AMOUNT = 1000  # $10.00
SUBSCRIPTION_INTERVAL = "month"

DB_PATH_DEFAULT = os.path.join(os.path.dirname(__file__), "facturae_b2b.db")


def get_db_connection(db_path: str = DB_PATH_DEFAULT) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn


def initialize_managed_payments_schema(db_path: str = DB_PATH_DEFAULT):
    """
    Initializes datastore tables and columns required for Managed Payments:
    - Products and default price cache
    - User subscription identifiers (customer_id, subscription_id, status)
    - Webhook event idempotency tracking
    """
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("PRAGMA synchronous=NORMAL;")

        # 1. Product and Price cache to reuse outputs from create_subscription_product
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS managed_payments_products (
                product_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                tax_code TEXT NOT NULL,
                default_price TEXT NOT NULL,
                currency TEXT NOT NULL,
                unit_amount INTEGER NOT NULL,
                created_at TEXT NOT NULL
            )
        """)

        # 2. Check and migrate user schema
        cursor.execute("PRAGMA table_info(usuarios)")
        columns = {col[1] for col in cursor.fetchall()}
        if "stripe_customer_id" not in columns:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN stripe_customer_id TEXT")
        if "stripe_subscription_id" not in columns:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN stripe_subscription_id TEXT")
        if "subscription_status" not in columns:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN subscription_status TEXT DEFAULT 'inactive'")
        if "subscription_period_end" not in columns:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN subscription_period_end TEXT")
        if "tipo_licencia" not in columns:
            cursor.execute("ALTER TABLE usuarios ADD COLUMN tipo_licencia TEXT")

        # 3. Transaction log & migration
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
        tx_columns = {col[1] for col in cursor.fetchall()}
        if "customer_id" not in tx_columns:
            cursor.execute("ALTER TABLE transacciones_stripe ADD COLUMN customer_id TEXT")
        if "subscription_id" not in tx_columns:
            cursor.execute("ALTER TABLE transacciones_stripe ADD COLUMN subscription_id TEXT")

        # 4. Webhook idempotency table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stripe_eventos_procesados (
                event_id TEXT PRIMARY KEY,
                event_type TEXT NOT NULL,
                processed_at TEXT NOT NULL,
                payload_summary TEXT
            )
        """)
        conn.commit()


# ==============================================================================
# STEP 1: CREATE ONE-TIME PRODUCT (Hamlet (e-book))
# ==============================================================================

def create_one_time_product(
    api_key: Optional[str] = None,
    name: str = ONE_TIME_PRODUCT_NAME,
    tax_code: str = ONE_TIME_TAX_CODE,
    currency: str = ONE_TIME_CURRENCY,
    unit_amount: int = ONE_TIME_UNIT_AMOUNT,
    db_path: str = DB_PATH_DEFAULT,
) -> Dict[str, Any]:
    """
    Step 1: create-one-time-product
    POST /v1/products with api_version 2026-02-25.preview
    Params:
      - name: "Hamlet (e-book)"
      - tax_code: "txcd_10103100"
      - default_price_data[currency]: "usd"
      - default_price_data[unit_amount]: 1000
    Returns the created product dictionary with default_price identifier.
    """
    key = api_key or os.environ.get("STRIPE_SECRET_KEY")
    if not key:
        raise ValueError("Stripe secret key must be provided or set in STRIPE_SECRET_KEY.")
    stripe.api_key = key

    # Check if product is already cached locally
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT product_id, default_price FROM managed_payments_products WHERE name = ? LIMIT 1",
            (name,),
        )
        cached = cursor.fetchone()
        if cached and cached["default_price"]:
            logger.info(f"Reusing existing cached product {cached['product_id']} with price {cached['default_price']}")
            return {
                "id": cached["product_id"],
                "default_price": cached["default_price"],
                "name": name,
            }

    # API Request to /v1/products
    logger.info("Executing API request POST /v1/products (api_version: 2026-02-25.preview)")
    product = stripe.Product.create(
        name=name,
        tax_code=tax_code,
        default_price_data={
            "currency": currency.lower(),
            "unit_amount": unit_amount,
        },
        stripe_version=STRIPE_API_VERSION_PREVIEW,
    )

    default_price = product.default_price
    default_price_id = default_price.id if hasattr(default_price, "id") else str(default_price)

    # Persist in local datastore
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT OR REPLACE INTO managed_payments_products
            (product_id, name, tax_code, default_price, currency, unit_amount, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product.id,
            name,
            tax_code,
            default_price_id,
            currency.lower(),
            unit_amount,
            ahora,
        ))
        conn.commit()

    logger.info(f"Created one-time product: {product.id} with default_price: {default_price_id}")
    return {
        "id": product.id,
        "default_price": default_price_id,
        "name": product.name,
    }


# ==============================================================================
# STEP 1 (VARIANT): CREATE SUBSCRIPTION PRODUCT
# ==============================================================================

def create_subscription_product(
    api_key: Optional[str] = None,
    db_path: str = DB_PATH_DEFAULT,
) -> Dict[str, Any]:
    """
    Step 1: create-subscription-product
    POST /v1/products with api_version 2026-02-25.preview
    Params:
      - name: "Basic subscription"
      - tax_code: "txcd_10103100"
      - default_price_data[currency]: "usd"
      - default_price_data[unit_amount]: 1000
      - default_price_data[recurring][interval]: "month"
    Returns the created product dictionary with default_price identifier.
    """
    key = api_key or os.environ.get("STRIPE_SECRET_KEY")
    if not key:
        raise ValueError("Stripe secret key must be provided or set in STRIPE_SECRET_KEY.")
    stripe.api_key = key

    # Check if product is already cached locally
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT product_id, default_price FROM managed_payments_products WHERE name = ? LIMIT 1",
            (SUBSCRIPTION_PRODUCT_NAME,),
        )
        cached = cursor.fetchone()
        if cached and cached["default_price"]:
            logger.info(f"Reusing existing cached product {cached['product_id']} with price {cached['default_price']}")
            return {
                "id": cached["product_id"],
                "default_price": cached["default_price"],
                "name": SUBSCRIPTION_PRODUCT_NAME,
            }

    # API Request to /v1/products
    logger.info("Executing API request POST /v1/products (api_version: 2026-02-25.preview)")
    product = stripe.Product.create(
        name=SUBSCRIPTION_PRODUCT_NAME,
        tax_code=SUBSCRIPTION_TAX_CODE,
        default_price_data={
            "currency": SUBSCRIPTION_CURRENCY,
            "unit_amount": SUBSCRIPTION_UNIT_AMOUNT,
            "recurring": {
                "interval": SUBSCRIPTION_INTERVAL,
            },
        },
        stripe_version=STRIPE_API_VERSION_PREVIEW,
    )

    default_price = product.default_price
    if hasattr(default_price, "id"):
        default_price_id = default_price.id
    else:
        default_price_id = str(default_price)

    # Persist in local datastore
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT OR REPLACE INTO managed_payments_products
            (product_id, name, tax_code, default_price, currency, unit_amount, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product.id,
            SUBSCRIPTION_PRODUCT_NAME,
            SUBSCRIPTION_TAX_CODE,
            default_price_id,
            SUBSCRIPTION_CURRENCY,
            SUBSCRIPTION_UNIT_AMOUNT,
            ahora,
        ))
        conn.commit()

    logger.info(f"Created subscription product: {product.id} with default_price: {default_price_id}")
    return {
        "id": product.id,
        "default_price": default_price_id,
        "name": product.name,
    }


# ==============================================================================
# STEP 2: CREATE CHECKOUT SESSION
# ==============================================================================

def create_checkout_session(
    default_price: str,
    success_url: str,
    mode: str = "payment",
    cancel_url: Optional[str] = None,
    customer_id: Optional[str] = None,
    customer_email: Optional[str] = None,
    client_reference_id: Optional[str] = None,
    api_key: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Step 2: create-checkout-session
    POST /v1/checkout/sessions with api_version 2026-02-25.preview
    Params:
      - line_items[0][price]: default_price (from Step 1)
      - line_items[0][quantity]: 1
      - managed_payments[enabled]: True
      - mode: "payment" (or "subscription")
      - success_url: application_url
    Returns the Checkout Session object with .url (for browser interaction) and .id.
    """
    key = api_key or os.environ.get("STRIPE_SECRET_KEY")
    if not key:
        raise ValueError("Stripe secret key must be provided or set in STRIPE_SECRET_KEY.")
    stripe.api_key = key

    params: Dict[str, Any] = {
        "line_items": [{
            "price": default_price,
            "quantity": 1,
        }],
        "mode": mode,
        "managed_payments": {
            "enabled": True,
        },
        "success_url": success_url,
    }

    if cancel_url:
        params["cancel_url"] = cancel_url
    if customer_id:
        params["customer"] = customer_id
    elif customer_email:
        params["customer_email"] = customer_email.strip().lower()
    if client_reference_id:
        params["client_reference_id"] = client_reference_id.strip()

    logger.info("Executing API request POST /v1/checkout/sessions (api_version: 2026-02-25.preview)")
    session = stripe.checkout.Session.create(
        **params,
        stripe_version=STRIPE_API_VERSION_PREVIEW,
    )

    logger.info(f"Created checkout session: {session.id}, url: {session.url}")
    return {
        "id": session.id,
        "url": session.url,
        "customer": session.customer,
        "subscription": session.subscription,
    }


# ==============================================================================
# STEP 4: HANDLE WEBHOOKS (checkout.session.completed)
# ==============================================================================

def handle_checkout_session_completed(
    session_data: Dict[str, Any],
    event_id: Optional[str] = None,
    db_path: str = DB_PATH_DEFAULT,
) -> Tuple[bool, str]:
    """
    Step 4: wait-for-checkout-completed / handle checkout.session.completed
    Processes the completed checkout session snapshot:
      - Reuses session.id, customer_id, subscription_id
      - Persists identifiers and updates domain models (user access, transactions)
      - Ensures idempotency
    """
    session_id = session_data.get("id")
    customer_id = session_data.get("customer")
    subscription_id = session_data.get("subscription")
    email = (
        session_data.get("client_reference_id")
        or session_data.get("customer_email")
        or session_data.get("customer_details", {}).get("email")
        or ""
    ).strip().lower()
    amount_total = session_data.get("amount_total", 1000)
    currency = session_data.get("currency", "usd")

    if not session_id:
        return False, "Invalid session data: missing session id."

    with get_db_connection(db_path) as conn:
        cursor = conn.cursor()

        # Check idempotency if event_id is provided
        if event_id:
            cursor.execute("SELECT event_id FROM stripe_eventos_procesados WHERE event_id = ?", (event_id,))
            if cursor.fetchone():
                return True, f"Event {event_id} already processed (idempotent skip)."

        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mode = session_data.get("mode")
        is_sub = (mode == "subscription") or bool(subscription_id)
        tipo_licencia = "mensual" if is_sub else "vitalicia"

        # 1. Update user domain model
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
        elif customer_id:
            cursor.execute("""
                UPDATE usuarios
                SET es_pro = 1,
                    tipo_licencia = ?,
                    fecha_pro = COALESCE(fecha_pro, ?),
                    stripe_subscription_id = COALESCE(?, stripe_subscription_id),
                    subscription_status = 'active'
                WHERE stripe_customer_id = ?
            """, (tipo_licencia, ahora, subscription_id, customer_id))

        # 2. Record transaction
        cursor.execute("""
            INSERT OR REPLACE INTO transacciones_stripe
            (session_id, email, monto_cents, moneda, fecha_pago, customer_id, subscription_id)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (session_id, email or "managed_payments_customer", amount_total, currency, ahora, customer_id, subscription_id))

        # 3. Record in idempotency log
        if event_id:
            cursor.execute("""
                INSERT INTO stripe_eventos_procesados (event_id, event_type, processed_at, payload_summary)
                VALUES (?, 'checkout.session.completed', ?, ?)
            """, (event_id, ahora, f"session_id={session_id}"))

        conn.commit()

    return True, f"Checkout session {session_id} successfully fulfilled and user activated."
