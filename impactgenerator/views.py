from django.shortcuts import render
from django.http import HttpResponse


def home(request):
    return render(request, 'impactgenerator/home.html', context={})


def upload(request):
    return render(request, 'impactgenerator/upload.html', context={})
