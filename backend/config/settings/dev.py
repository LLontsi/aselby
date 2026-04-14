from .base import *

DEBUG = True
ALLOWED_HOSTS = ['*']

# Storage simple pour le dev — pas besoin de collectstatic
STATICFILES_STORAGE = 'django.contrib.staticfiles.storage.StaticFilesStorage'

# Debug toolbar désactivé pour éviter les conflits de version
# Pour l'activer : pip install django-debug-toolbar et décommenter les lignes ci-dessous
# INSTALLED_APPS += ['debug_toolbar']
# MIDDLEWARE = ['debug_toolbar.middleware.DebugToolbarMiddleware'] + MIDDLEWARE
# INTERNAL_IPS = ['127.0.0.1']
