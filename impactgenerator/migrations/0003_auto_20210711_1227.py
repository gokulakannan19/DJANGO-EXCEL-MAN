# Generated by Django 3.2.5 on 2021-07-11 06:57

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('impactgenerator', '0002_auto_20210710_1656'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='document',
            name='csv',
        ),
        migrations.RemoveField(
            model_name='document',
            name='read_excel',
        ),
        migrations.RemoveField(
            model_name='document',
            name='write_excel',
        ),
        migrations.AddField(
            model_name='document',
            name='document',
            field=models.FileField(blank=True, null=True, upload_to='csv'),
        ),
    ]
