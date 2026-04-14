from django.urls import path
from . import views
app_name = 'banque'
urlpatterns = [
    path('', views.historique, name='historique'),
    path('cheques/', views.cheques, name='cheques'),
    path('tresorerie/', views.tresorerie, name='tresorerie'),
]