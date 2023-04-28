from django.db import models
from datetime import datetime

class Products(models.Model):

    COLOR =(
        ("WHITE", 'White'),
        ("BLUE", 'Blue'),
        ("BLACK", 'Black'),
        ("GREEN", 'Green'),
    )
    name = models.CharField(max_length=100)
    color = models.CharField(max_length=10, choices=COLOR)
    price = models.IntegerField()

    def __str__(self):
        return '{}: - {}: - {}:'.format(self.name, self.color, self.price)

class DataPoint(models.Model):
    x = models.IntegerField()
    y = models.IntegerField()   

class DataPointTime(models.Model):
    x = models.DateTimeField(default=datetime.now)
    y = models.IntegerField()   