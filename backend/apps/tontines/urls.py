from django.urls import path
from . import views
app_name = 'tontines'
urlpatterns = [
    path('',                        views.tableau_tontines,    name='tableau'),
    path('telecharger/<str:niveau>/', views.telecharger_tontine, name='telecharger'),
]