from rest_framework import serializers
from .models import Tipo, Rol, Usuario, Area, Etapa, Ambiente, Proyecto, Formulario, Detalle_Formulario, Bitacora, Workspace, Programa, Version_Cobol, Region_Cics, Tipo_Programa, Tipo_Formulario

class TipoSerializer(serializers.ModelSerializer):
    class Meta:
        model = Tipo
        fields = '__all__'

class RolSerializer(serializers.ModelSerializer):
    class Meta:
        model = Rol
        fields = '__all__'

class UsuaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Usuario
        fields = '__all__'

class AreaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Area
        fields = '__all__'

class EtapSerializer(serializers.ModelSerializer):
    class Meta:
        model = Etapa
        fields = '__all__'

class AmbiSerializer(serializers.ModelSerializer):
    class Meta:
        model = Ambiente
        fields = '__all__'

class ProySerializer(serializers.ModelSerializer):
    class Meta:
        model = Proyecto
        fields = '__all__'

class FormSerializer(serializers.ModelSerializer):
    class Meta:
        model = Formulario
        fields = '__all__'

class DtfmSerializer(serializers.ModelSerializer):
    class Meta:
        model = Detalle_Formulario
        fields = '__all__'

class BitaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Bitacora
        fields = '__all__'

class WorkSerializer(serializers.ModelSerializer):
    class Meta:
        model = Workspace
        fields = '__all__'

class ProgSerializer(serializers.ModelSerializer):
    class Meta:
        model = Programa
        fields = '__all__'

class VersCobSerializer(serializers.ModelSerializer):
    class Meta:
        model = Version_Cobol
        fields = '__all__'

class RegCSerializer(serializers.ModelSerializer):
    class Meta:
        model = Region_Cics
        fields = '__all__'

class TipProgSerializer(serializers.ModelSerializer):
    class Meta:
        model = Tipo_Programa
        fields = '__all__'

class TipFormSerializer(serializers.ModelSerializer):
    class Meta:
        model = Tipo_Formulario
        fields = '__all__'