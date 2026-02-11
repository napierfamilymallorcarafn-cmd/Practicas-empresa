from django.urls import path, include
from django.conf.urls.static import static
from django.conf import settings
from . import views

# Archivo que define las urls de la aplicación, y redirigue las peticiones del usuario a views concretas
urlpatterns = [
    path('', views.muestras_todas, name='principal'),
    path('muestras/', views.muestras_todas, name='muestras_todas'),
    path('muestras/nuevo_centro', views.nuevo_centro, name='nuevo_centro'),
    path('muestras/acciones_post', views.acciones_post, name='acciones_post'),
    path('muestras/acciones_post/seleccionar_estudio', views.seleccionar_estudio, name='seleccionar_estudio'),
    path('muestras/acciones_post/seleccionar_estudio/añadir_muestras_estudio',views.añadir_muestras_estudio,name='añadir_muestras_estudio'),
    path('muestras/acciones_post/cambio_posicion', views.cambio_posicion, name='cambio_posicion'),
    path('muestras/nueva', views.añadir_muestras, name='añadir_muestras'),
    path('muestras/historial_localizaciones/<int:muestra_id>', views.historial_localizaciones_muestra, name='historial_localizaciones'),
    path('muestras/envio/agenda', views.agenda, name='agenda'),
    path('muestras/envio/agenda/editar/<int:id_centro>',views.editar_centro,name='editar_centro'),
    path('muestras/envio/agenda/<int:centro>/envio', views.formulario_envios, name='formulario_envio'),
    path('muestras/envio/agenda/<int:centro>/envio/registrar_envio', views.registrar_envio, name='registrar_envio'),
    path('muestras/envio/agenda/<int:centro>/envio/upload_excel_envio', views.upload_excel_envios, name='upload_excel_envios'),
    path('muestras/envio/agenda/eliminar_centro', views.eliminar_centro, name="eliminar_centro"),
    path('muestras/historial_envios/<int:muestra_id>', views.historial_envios, name='historial_envios'),
    path('muestras/historial_estudios/<int:muestra_id>', views.historial_estudios_muestra, name='historial_estudios'),
    path('muestras/upload_excel', views.upload_excel, name='upload_excel'),
    path('muestras/upload_excel/descargar/<int:macro>', views.descargar_plantilla, name='descargar_plantilla_muestras'),
    path('archivo/detalles_muestra/<str:nom_lab>', views.detalles_muestra, name='detalles_muestra'),
    path('muestras/detalles_muestra/<str:nom_lab>/editar', views.editar_muestra, name='editar_muestra'),
    path('muestras/detalles_muestra/<str:nom_lab>/eliminar', views.eliminar_muestra, name='eliminar_muestra'),
    path('archivo/', views.localizaciones, name='localizaciones_todas'),
    path('archivo/detalles_congelador/<str:nombre_congelador>', views.detalles_congelador, name ='detalles_congelador'),
    path('archivo/detalles_congelador/<str:nombre_congelador>/editar', views.editar_congelador, name ='editar_congelador'),
    path('archivo/nuevo', views.upload_excel_localizaciones, name='upload_excel_localizaciones'),
    path('archivo/nuevo/<int:macro>', views.descargar_plantilla, name='descargar_plantilla_localizaciones'),
    path('archivo/nuevo/<int:macro>', views.descargar_plantilla, name='descargar_plantilla_localizaciones_macros'),
    path('archivo/eliminar_localizacion', views.eliminar_localizacion, name="eliminar_localizacion"),
    path('estudios/',views.estudios_todos, name='estudios_todos'),
    path('estudios/excel',views.excel_estudios, name='excel_estudios'),
    path('estudios/nuevo',views.nuevo_estudio, name='nuevo_estudio'),
    path('estudios/<int:id_estudio>', views.repositorio_estudio, name="repositorio_estudio"),
    path('estudios/<int:id_estudio>/subir', views.subir_documento, name="subir_documento"),
    path('estudios/<int:id_estudio>/editar', views.editar_estudio, name="editar_estudio"),
    path('estudios/<int:id_estudio>/eliminar', views.eliminar_estudio, name="eliminar_estudio"),
    path('estudios/<int:id_estudio>/<int:documento_id>',views.descargar_documento, name="descargar_documento"),
    path('estudios/eliminar_documento',views.eliminar_documento, name="eliminar_documento"),
    # Rutas AJAX para cargar dinámicamente las localizaciones
    path('api/get_estantes_por_congelador/', views.get_estantes_por_congelador, name='get_estantes_por_congelador'),
    path('api/get_racks_por_estante/', views.get_racks_por_estante, name='get_racks_por_estante'),
    path('api/get_cajas_por_rack/', views.get_cajas_por_rack, name='get_cajas_por_rack'),
    path('api/get_subposiciones_por_caja/', views.get_subposiciones_por_caja, name='get_subposiciones_por_caja'),
    path('api/get_subposiciones_por_caja_tree/', views.get_subposiciones_por_caja_tree, name='get_subposiciones_por_caja_tree'),
]

# Ajuste para poder servir los archivos estáticos cuando el DEBUG es True
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)