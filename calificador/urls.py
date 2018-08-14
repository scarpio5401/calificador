from django.conf.urls import patterns, include, url
from django.contrib import admin


urlpatterns = patterns('',
    # Examples:
    url(r'^grader/', include('grader.urls')),
    url(r'^admin/', include(admin.site.urls)),
)
