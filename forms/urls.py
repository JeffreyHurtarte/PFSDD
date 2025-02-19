from django.urls import path, include
from rest_framework.documentation import include_docs_urls
from rest_framework import routers
from forms import views
from .views import generar_traslado_fases, generar_solicitud_programas, excel_a_pdf

router = routers.DefaultRouter()

router.register(r'tipos', views.TipoView, 'tipos')
router.register(r'roles', views.RolView, 'roles')
router.register(r'usuarios', views.UsuaView, 'usuarios')
router.register(r'areas', views.AreaView, 'areas')
router.register(r'etapas', views.EtapView, 'etapas')
router.register(r'ambientes', views.AmbiView, 'ambientes')
router.register(r'proyectos', views.ProyView, 'proyectos')
router.register(r'formularios', views.FormView, 'formularios')
router.register(r'detalle_forms', views.DtfmView, 'detalle_forms')
router.register(r'bitacora', views.BitaView, 'bitacora')
router.register(r'corporativo', views.UsuaCorpView, 'corporativo')
router.register(r'workspaces', views.WorkView, 'workspaces')
router.register(r'programas', views.ProgView, 'programas')
router.register(r'version_cobol', views.VersCobView, 'version_cobol')
router.register(r'region_cics', views.RegCView, 'region_cics')
router.register(r'tipo_programa', views.TipProgView, 'tipo_programa')
router.register(r'tipo_formulario', views.TipFormView, 'tipo_formulario')
router.register(r'login', views.LoginView, 'login')


urlpatterns = [
    path("api/v1/", include(router.urls)),
    path('docs/', include_docs_urls(title="Forms API"))
] 

urlpatterns += [
    path('generar_traslado_fases/', generar_traslado_fases, name='generar_traslado_fases'),
    path('generar_solicitud_programas/', generar_solicitud_programas, name='generar_solicitud_programas'),
    path('excel-a-pdf/', excel_a_pdf, name='excel_a_pdf')
]
