from django.conf.urls import include, url
from views import *

urlpatterns = [
    url(r'index/$', index),
	url(r'excel/(?P<token>\d+)/$', exportar_excel),
]