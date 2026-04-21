import os
from .base import *

DEBUG = False

# ── Hosts & CSRF ───────────────────────────────────────────────
ALLOWED_HOSTS = os.environ.get('ALLOWED_HOSTS', 'localhost').split(',')

CSRF_TRUSTED_ORIGINS = [
    origin.strip()
    for origin in os.environ.get('CSRF_TRUSTED_ORIGINS', '').split(',')
    if origin.strip()
]

# ── Sécurité ───────────────────────────────────────────────────
STATICFILES_STORAGE      = 'whitenoise.storage.CompressedManifestStaticFilesStorage'
SECURE_BROWSER_XSS_FILTER = True
SECURE_CONTENT_TYPE_NOSNIFF = True
X_FRAME_OPTIONS          = 'DENY'
SESSION_COOKIE_SECURE    = True
CSRF_COOKIE_SECURE       = True
