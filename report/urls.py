from django.urls import path
from .views import home_view, upload_file_view


urlpatterns = [
    path('', home_view, name = 'home'),
    # переход на upload/ и выполнение функции загрузки файла
    path('upload/', upload_file_view, name='upload_file'),
]