# Generated by Django 5.0.7 on 2024-10-22 06:13

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('forms', '0017_remove_usuario_is_active_remove_usuario_is_staff_and_more'),
    ]

    operations = [
        migrations.CreateModel(
            name='Region_Cics',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nombre', models.CharField(max_length=10)),
            ],
        ),
        migrations.CreateModel(
            name='Tipo_Programa',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nombre', models.CharField(max_length=10)),
            ],
        ),
        migrations.CreateModel(
            name='Version_Cobol',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nombre', models.CharField(max_length=10)),
            ],
        ),
        migrations.CreateModel(
            name='Programa',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nombre', models.CharField(max_length=10)),
                ('version', models.IntegerField(default=0)),
                ('cics', models.IntegerField(default=0)),
                ('tipo_prog', models.IntegerField(default=0)),
                ('mapset_copy', models.CharField(max_length=10)),
                ('formulario', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='formulario_programa', to='forms.formulario')),
            ],
        ),
        migrations.CreateModel(
            name='Workspace',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nombre', models.CharField(max_length=500)),
                ('formulario', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='formulario_workspace', to='forms.formulario')),
            ],
        ),
    ]
