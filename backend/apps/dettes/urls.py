from django.urls import path
from . import views
app_name = 'dettes'
urlpatterns = [
    path('', views.liste_rouge, name='liste'),
    path('<int:pk>/', views.fiche_dette, name='fiche'),
]