# Generated by Django 3.2.6 on 2021-08-25 13:36

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0050_alter_bilan_agregat'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bilan',
            name='agregat',
            field=models.CharField(blank=True, default={'testform2', 'testform', 'testform3'}, max_length=255, null=True),
        ),
    ]
