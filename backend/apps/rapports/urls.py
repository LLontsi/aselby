from django.urls import path
from . import views
app_name = 'rapports'
urlpatterns = [path('', views.dashboard, name='dashboard')]