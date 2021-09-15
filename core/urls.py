from django.contrib import admin
from django.urls import path, include
from bilans import views
from django.conf.urls import url
from django.conf import settings
from django.conf.urls.static import static
from registration.backends.simple.views import RegistrationView



urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('bilans.urls', namespace='bilans')),
    path('accounts/', include('registration.backends.simple.urls')),
    path('', include('registration.backends.simple.urls')),
    
    
    
]  + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
 