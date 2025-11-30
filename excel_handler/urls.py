from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload_excel, name='upload_excel'),
    path('download/<int:file_id>/', views.download_processed, name='download_processed'),
    path('download-final/<int:file_id>/', views.download_final_file, name='download_final_file'),
    path('process-all/', views.process_all_workflows, name='process_all_workflows'),
    path('workflow4/', views.workflow4_view, name='workflow4'),
    path('chart-data/<int:file_id>/', views.get_chart_data, name='get_chart_data'),
    path('strict-extract/<int:file_id>/', views.strict_extract_excel, name='strict_extract_excel'),
]