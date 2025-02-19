from django.db import models

class Tipo(models.Model):
    nombre = models.CharField(max_length=50)

    def __str__(self):
        return self.nombre
    
class Rol(models.Model):
    nombre = models.CharField(max_length=20)

    def __str__(self):
        return self.nombre
    
class Usuario(models.Model):
    nombre = models.CharField(max_length=200)
    corporativo = models.IntegerField(default=0)
    antiguedad = models.IntegerField(default=0)
    usuario_tso = models.CharField(max_length=4)
    contrasena =  models.CharField(max_length=256, default='1234')
    rol = models.IntegerField(default=0)
    firma = models.BooleanField(default=False)

    def __str__(self):
        return str(self.corporativo)
    
class Area(models.Model):
    nombre = models.CharField(max_length=100)

    def __str__(self):
        return self.nombre
    
class Etapa(models.Model):
    nombre = models.CharField(max_length=100)

    def __str__(self):
        return self.nombre
    
class Ambiente(models.Model):
    nombre = models.CharField(max_length=20)

    def __str__(self):
        return self.nombre

class Proyecto(models.Model):
    rational = models.IntegerField(default=0)
    nombre = models.CharField(max_length=200)
    tipo = models.IntegerField(default=0)
    usuario = models.IntegerField(default=0)
    coordinador = models.IntegerField(default=0)
    fecha = models.DateField(auto_now=True)
    area = models.IntegerField(default=0)
    descripcion = models.CharField(max_length=500)

    def __str__(self):
        return str(self.rational)

class Formulario(models.Model):
    nombre = models.CharField(max_length=100)
    proyecto = models.IntegerField(default=0)
    tipo = models.IntegerField(default=0)
    descripcion = models.CharField(max_length=250)

    def __str__(self):
        return self.nombre
    
class Detalle_Formulario(models.Model):
    fecha = models.DateTimeField(auto_now_add=True)
    formulario = models.IntegerField(default=0)
    ambiente = models.IntegerField(default=0)

    def __str__(self):
        return str(self.formulario)
    
class Workspace(models.Model):
    nombre = models.CharField(max_length=500)
    formulario_id = models.IntegerField(default=0)
    
    def __str__(self):
        return self.nombre

class Programa(models.Model):
    nombre = models.CharField(max_length=10)
    version = models.IntegerField(default=0)
    cics = models.IntegerField(default=0)
    tipo_prog = models.IntegerField(default=0)
    mapset_copy = models.CharField(max_length=10)
    formulario_id = models.IntegerField(default=0)
    
    def __str__(self):
        return self.nombre

class Version_Cobol(models.Model):
    nombre = models.CharField(max_length=10)
    
    def __str__(self):
        return self.nombre

class Region_Cics(models.Model):
    nombre = models.CharField(max_length=10)
    
    def __str__(self):
        return self.nombre

class Tipo_Programa(models.Model):
    nombre = models.CharField(max_length=10)
    
    def __str__(self):
        return self.nombre

class Tipo_Formulario(models.Model):
    nombre = models.CharField(max_length=50)
    
    def __str__(self):
        return self.nombre


class Bitacora(models.Model):
    proyecto = models.IntegerField(default=0)
    usuario = models.IntegerField(default=0)
    corporativo = models.IntegerField(default=0)
    evento = models.CharField(max_length=50)
    etapa_actual = models.IntegerField(default=0)
    etapa_traslado = models.IntegerField(default=0)
    fecha = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return str(self.proyecto)
    
class Mapas(models.Model):
    nombre = models.CharField(max_length=100)
    programa = models.IntegerField(default=0)

    def __str__(self):
        return self.nombres