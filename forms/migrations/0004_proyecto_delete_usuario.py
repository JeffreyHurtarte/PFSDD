# Generated by Django 5.0.7 on 2024-09-06 03:50

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('forms', '0003_rename_usuarios_usuario'),
    ]

    operations = [
        migrations.CreateModel(
            name='Proyecto',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('numero_rational', models.IntegerField(default=0)),
                ('nombre', models.CharField(max_length=200)),
                ('fecha_inicio', models.DateField(auto_now=True)),
                ('area', models.TextField(max_length=250)),
                ('coordinador', models.IntegerField(default=0)),
                ('descripcion', models.TextField(max_length=500)),
            ],
        ),
        migrations.DeleteModel(
            name='Usuario',
        ),
    ]
