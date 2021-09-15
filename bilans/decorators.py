from django.http import HttpResponse
from django.shortcuts import redirect
from django import template

def allowed_users(allowed_roles=[]):
    def decorator(view_func):
        def wrapper_func(request, *args, **kwargs):
            group = None
            if request.user.groups.exists():
                group = request.user.groups.all()[0].name
            
            if group in allowed_roles:
                return view_func(request, *args, **kwargs)
            
            else:
                return HttpResponse("vous n'etes pas autoriser a consulter cette page")
                
                    
            return view_func(request, *args, **kwargs)
        return wrapper_func
    return decorator




register = template.Library()

@register.filter(name='has_group')
def has_group(user, group_name):
    return user.groups.filter(name=group_name).exists()