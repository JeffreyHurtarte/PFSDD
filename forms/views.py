from rest_framework import viewsets
from rest_framework.response import Response
from .serializer import TipoSerializer, RolSerializer, UsuaSerializer, AreaSerializer, EtapSerializer, AmbiSerializer, ProySerializer, FormSerializer, DtfmSerializer, BitaSerializer, WorkSerializer, ProgSerializer, VersCobSerializer, RegCSerializer, TipProgSerializer, TipFormSerializer
from .models import Tipo, Rol, Usuario, Area, Etapa, Ambiente, Proyecto, Formulario, Detalle_Formulario, Bitacora, Workspace, Programa, Region_Cics, Tipo_Programa, Version_Cobol, Tipo_Formulario
from django.http import HttpResponse
from openpyxl import Workbook
from django.conf import settings
from .formatos import form_traslado_fases, form_solicitud_programas
import os
import pandas as pd
import pdfkit

def generar_traslado_fases(request):
    file_path = form_traslado_fases()
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename=traslado_fases.xlsx'
    with open(file_path, 'rb') as f:
        response.write(f.read())
    return response

def generar_solicitud_programas(request):
    file_path = form_solicitud_programas()
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename=solicitud_programas.xlsx'
    with open(file_path, 'rb') as f:
        response.write(f.read())
    return response

def excel_a_pdf(request):
    config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
    file_path = os.path.join(settings.MEDIA_ROOT, 'traslado_fases.xlsx')
    pdf_path = os.path.join(settings.MEDIA_ROOT, 'traslado_fases.pdf')
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    df = pd.read_excel(file_path)
    html = df.to_html()
    pdfkit.from_string(html, pdf_path, configuration=config)
    with open(pdf_path, 'rb') as pdf:
        response = HttpResponse(pdf.read(), content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename=traslado_fases.pdf'
        return response

class TipoView(viewsets.ModelViewSet):
    serializer_class = TipoSerializer 
    queryset = Tipo.objects.all()
    
class RolView(viewsets.ModelViewSet):
    serializer_class = RolSerializer
    queryset = Rol.objects.all()

class UsuaView(viewsets.ModelViewSet):
    serializer_class = UsuaSerializer
    queryset = Usuario.objects.all()

class AreaView(viewsets.ModelViewSet):
    serializer_class = AreaSerializer 
    queryset = Area.objects.all()

class EtapView(viewsets.ModelViewSet):
    serializer_class = EtapSerializer
    queryset = Etapa.objects.all()

class AmbiView(viewsets.ModelViewSet):
    serializer_class = AmbiSerializer
    queryset = Ambiente.objects.all()

class ProyectoView(viewsets.ModelViewSet):
    serializer_class = ProySerializer 
    queryset = Proyecto.objects.all()

class FormView(viewsets.ModelViewSet):
    serializer_class = FormSerializer
    queryset = Formulario.objects.all()

    def get_queryset(self):
        queryset = super().get_queryset()
        proyecto = self.request.query_params.get('proyecto')
        if proyecto is not None:
            queryset = queryset.filter(proyecto=proyecto)
        return queryset

class DtfmView(viewsets.ModelViewSet):
    serializer_class = DtfmSerializer 
    queryset = Detalle_Formulario.objects.all()

class BitaView(viewsets.ModelViewSet):
    serializer_class = BitaSerializer
    queryset = Bitacora.objects.all()

class WorkView(viewsets.ModelViewSet):
    serializer_class = WorkSerializer
    queryset = Workspace.objects.all()

class ProgView(viewsets.ModelViewSet):
    serializer_class = ProgSerializer
    queryset = Programa.objects.all()

class VersCobView(viewsets.ModelViewSet):
    serializer_class = VersCobSerializer
    queryset = Version_Cobol.objects.all()

class RegCView(viewsets.ModelViewSet):
    serializer_class = RegCSerializer
    queryset = Region_Cics.objects.all()

class TipProgView(viewsets.ModelViewSet):
    serializer_class = TipProgSerializer
    queryset = Tipo_Programa.objects.all()

class TipFormView(viewsets.ModelViewSet):
    serializer_class = TipFormSerializer
    queryset = Tipo_Formulario.objects.all()    

class ProyView(viewsets.ModelViewSet):
    queryset = Proyecto.objects.all()
    serializer_class = ProySerializer

    def list(self, request, *args, **kwargs):
        corporativo = request.query_params.get('corporativo', None)
        if corporativo is not None:
            try:
                usuario = Usuario.objects.get(corporativo=corporativo)
                proyectos = Proyecto.objects.filter(usuario=usuario.id)
                serializer = self.get_serializer(proyectos, many=True)
                return Response(serializer.data)
            except Usuario.DoesNotExist:
                return Response({"error": "Usuario no encontrado"}, status=404)
        return super().list(request, *args, **kwargs)
    
class UsuaCorpView(viewsets.ModelViewSet):
    queryset = Usuario.objects.all()
    serializer_class = UsuaSerializer

    def list(self, request, *args, **kwargs):
        corporativo = request.query_params.get('corporativo', None)
        if corporativo is not None:
            try:
                usuario = Usuario.objects.get(corporativo=corporativo)
                serializer = self.get_serializer(usuario)
                return Response(serializer.data)
            except Usuario.DoesNotExist:
                return Response({"error": "Usuario no encontrado"}, status=404)
        return super().list(request, *args, **kwargs)
    
class LoginView(viewsets.ViewSet):
    serializer_class = UsuaSerializer

    def create(self, request):
        corporativo = request.data.get('corporativo')
        try:
            usuario = Usuario.objects.get(corporativo=corporativo)
            serializer = UsuaSerializer(usuario)
            return Response(serializer.data, status=200)
        except Usuario.DoesNotExist:
            return Response({'error': 'Usuario no encontrado'}, status=404)