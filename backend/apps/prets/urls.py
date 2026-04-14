from django.urls import path
from . import views
app_name = 'prets'
urlpatterns = [
    path('', views.liste_prets, name='liste'),
    path('nouveau/', views.nouveau_pret, name='nouveau'),
    path('rapport-retards/', views.rapport_retards, name='rapport_retards'),
    path('<int:pk>/', views.detail_pret, name='detail'),
    path('<int:pk>/valider/', views.valider_demande, name='valider'),
    path('<int:pk>/rejeter/', views.rejeter_demande, name='rejeter'),
]