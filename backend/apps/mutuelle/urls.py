from django.urls import path
from . import views
app_name = 'mutuelle'
urlpatterns = [
    path('', views.etat_mutuelle, name='etat'),
    path('aide/nouvelle/', views.nouvelle_aide, name='nouvelle_aide'),
]