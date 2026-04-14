from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('apps.public.urls')),
    path('', include('apps.users.urls')),
    path('mon-espace/', include(('apps.users.membre_urls', 'membre'))),
    path('dashboard/', include('apps.rapports.urls')),
    path('parametrage/', include('apps.parametrage.urls')),
    path('adherents/', include('apps.adherents.urls')),
    path('tontines/', include('apps.tontines.urls')),
    path('fonds/', include('apps.fonds.urls')),
    path('saisie/', include('apps.saisie.urls')),
    path('prets/', include('apps.prets.urls')),
    path('mutuelle/', include('apps.mutuelle.urls')),
    path('foyer/', include('apps.foyer.urls')),
    path('banque/', include('apps.banque.urls')),
    path('exercice/', include('apps.exercice.urls')),
    path('dettes/', include('apps.dettes.urls')),
]

handler403 = 'apps.core.views.erreur_403'
handler404 = 'apps.core.views.erreur_404'
handler500 = 'apps.core.views.erreur_500'

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    if 'debug_toolbar' in settings.INSTALLED_APPS:
        import debug_toolbar
        urlpatterns += [path('__debug__/', include(debug_toolbar.urls))]
