from django.db import models


class Entry(models.Model):
    year = models.CharField(max_length=10, default="")
    month = models.CharField(max_length=20, default="")
    category = models.CharField(max_length=10, default="")
    clubbed_name = models.CharField(max_length=100, default="")
    product = models.CharField(max_length=100, default="")
    value = models.CharField(max_length=100, default="")

    def __str__(self):
        return self.clubbed_name