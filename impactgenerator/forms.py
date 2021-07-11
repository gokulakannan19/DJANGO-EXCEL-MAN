from django import forms
from django.forms import fields

from .models import Document


class DocumentForm(forms.ModelForm):
    class Meta:
        model = Document
        fields = ('name', 'document')
