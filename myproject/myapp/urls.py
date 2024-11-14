from django.urls import path
from .views import generate_excel

urlpatterns = [
    path('', generate_excel, name='generate_excel'),
]