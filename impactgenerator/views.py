from django import forms
from django.shortcuts import render, redirect
from django.core.files.storage import FileSystemStorage
from .forms import DocumentForm
# import pandas as pd
# import openpyxl
# from openpyxl import Workbook
# from openpyxl.styles import Alignment


def home(request):
    if request.method == "POST":
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return redirect('home')
    else:
        form = DocumentForm()
        context = {
            'form': form,
        }

    return render(request, 'impactgenerator/home.html', context)


# def upload(request):
#     if request.method == "POST":
#         uploaded_file = request.FILES['document']
#         fs = FileSystemStorage()
#         fs.save(uploaded_file.name, uploaded_file)
#         print(uploaded_file.name)
#         print(uploaded_file.size)
#     return render(request, 'impactgenerator/upload.html', context={})


# def document(request):
#     return render(request, 'impactgenerator/document.html', context={})


# def upload_document(request):

#     if request.method == "POST":
#         form = DocumentForm(request.POST, request.FILES)
#         if form.is_valid():
#             form.save()
#             return redirect('document')
#     else:
#         form = DocumentForm()
#         context = {
#             'form': form,
#         }

#     return render(request, 'impactgenerator/upload_document.html', context)
