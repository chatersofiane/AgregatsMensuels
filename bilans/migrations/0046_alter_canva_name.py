# Generated by Django 3.2.6 on 2021-08-22 09:45

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0045_auto_20210822_1039'),
    ]

    operations = [
        migrations.AlterField(
            model_name='canva',
            name='name',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
    ]