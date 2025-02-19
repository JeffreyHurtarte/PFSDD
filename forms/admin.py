from django.contrib import admin
from .models import Tipo, Rol, Usuario, Area, Etapa, Ambiente, Proyecto, Formulario, Detalle_Formulario, Bitacora

# Register your models here.

admin.site.register(Tipo)
admin.site.register(Rol)
admin.site.register(Usuario)
admin.site.register(Area)
admin.site.register(Etapa)
admin.site.register(Ambiente)
admin.site.register(Proyecto)
admin.site.register(Formulario)
admin.site.register(Detalle_Formulario)
admin.site.register(Bitacora)