from django.db import models

# Create your models here.


class Document(models.Model):
    name = models.CharField(max_length=200, null=True)
    csv = models.FileField(upload_to="document/csv")
    date_created = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name
