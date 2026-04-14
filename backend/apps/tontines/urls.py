from django.urls import path
from . import views
app_name = 'tontines'
urlpatterns = [
    path('', views.tableau_mensuel, name='tableau'),
    path('detail/<str:niveau_code>/', views.detail_participants, name='detail_participants'),
    path('enchere/', views.enchere, name='enchere'),
    path('calendrier/', views.calendrier, name='calendrier'),
    path('interets/', views.repartition_interets, name='interets'),
]