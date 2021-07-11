from django.db import models

# Create your models here.


class Document(models.Model):
    name = models.CharField(max_length=200, null=True)
    document = models.FileField(upload_to="csv", null=True, blank=True)
    # read_excel = models.FileField(
    #     upload_to="read_excel", null=True,  blank=True)
    # write_excel = models.FileField(
    #     upload_to="write_excel", null=True,  blank=True)
    date_created = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name
