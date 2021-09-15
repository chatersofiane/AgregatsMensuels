from django.urls import path
from . import views
from django.contrib.auth.decorators import login_required


app_name = 'bilans'

urlpatterns = [
    path('', views.HomeView.as_view(), name='home'),
    path('canvas/', login_required(views.index), name='list_canvas'),
    path('canvas/add/', login_required(views.CanvaCreateView.as_view()), name='add_canvas'),
    path('canvas/edit/<int:pk>/', login_required(views.CanvaDetailView.as_view()), name='canva_detail'),
    path('canvas/<int:pk>/bilan/edit/', login_required(views.BilansEditView.as_view()), name='bilan_edit'),
    path('canvas/<int:pk>/autre/edit/', login_required(views.AutresEditView.as_view()), name='autre_edit'),
    path('canvas/<int:pk>/tresorerie/edit/', login_required(views.TresoreriesEditView.as_view()), name='tresorerie_edit'),
    path('canvas/<int:pk>/production/edit/', login_required(views.ProductionsEditView.as_view()), name='production_edit'),
    path('canvas/<int:pk>/validation/edit/', login_required(views.CheckboxsEditView.as_view()), name='canva_valid'),
    path('canva/excelpage/', views.ExcelPageView.as_view(), name='excelpage'), 
   
    path('canvas/edit/bilanexport/<int:pk>', views.export_bilans_xls, name='bilan_excel'),
    path('canvas/consult/sommeexport/<int:pk>', views.export_somme_xls, name='somme_excel'),
    
    
    
    path('canvas/consult/<int:pk>/', login_required(views.CanvaConsultlView.as_view()), name='canva2_detail'),
   
    
   path('canvas/list2/', login_required(views.index2), name='list2_canvas'),
    
    
  


    
    
            
]

