# Generated by Django 5.0.7 on 2024-10-20 01:08

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('forms', '0013_ambiente_area_bitacora_detalle_formulario_rol_tipo_and_more'),
    ]

    operations = [
        migrations.AlterField(
            model_name='usuario',
            name='contrasena',
            field=models.CharField(max_length=256),
        ),
    ]
