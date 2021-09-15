from django.contrib import admin

from .models import Canva, Bilan, Checkbox



class canvaAdmin(admin.ModelAdmin):
    list_display = ('name', 'site', 'mois', 'année', 'created')
    

    
    




admin.site.register(Canva, canvaAdmin)
admin.site.register(Bilan)
admin.site.register(Checkbox)