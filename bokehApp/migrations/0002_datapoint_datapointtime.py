# Generated by Django 4.1.4 on 2023-01-04 15:38

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bokehApp', '0001_initial'),
    ]

    operations = [
        migrations.CreateModel(
            name='DataPoint',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('x', models.IntegerField()),
                ('y', models.IntegerField()),
            ],
        ),
        migrations.CreateModel(
            name='DataPointTime',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('x', models.DateTimeField(default=datetime.datetime.now)),
                ('y', models.IntegerField()),
            ],
        ),
    ]