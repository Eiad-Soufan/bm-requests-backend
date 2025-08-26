from django.db import models
from django.contrib.auth.models import AbstractUser
from django.conf import settings

class CustomUser(AbstractUser):
    ROLE_CHOICES = [
        ('manager', 'Management'),
        ('hr', 'HR'),
        ('employee', 'Employee'),
    ]
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='employee')

    def __str__(self):
        return self.username


class Section(models.Model):
    name_ar = models.CharField(max_length=100)
    name_en = models.CharField(max_length=100)

    def __str__(self):
        return self.name_ar

class FormModel(models.Model):
    section = models.ForeignKey(Section, on_delete=models.CASCADE, related_name='forms')
    serial_number = models.CharField(max_length=20, unique=True)
    name_ar = models.CharField(max_length=100)
    name_en = models.CharField(max_length=100)
    category = models.CharField(max_length=100)
    description = models.TextField(blank=True)
    file = models.FileField(upload_to='forms/') 

    def __str__(self):
        return f"{self.name_ar} ({self.serial_number})"

class UserSectionPermission(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    section = models.ForeignKey(Section, on_delete=models.CASCADE)

    class Meta:
        unique_together = ('user', 'section')


class Notification(models.Model):
    IMPORTANCE_CHOICES = [
        ('normal', 'عادي'),
        ('important', 'هام'),
    ]
    title = models.CharField(max_length=255)
    message = models.TextField()
    importance = models.CharField(max_length=10, choices=IMPORTANCE_CHOICES, default='normal')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.title

class UserNotification(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    notification = models.ForeignKey(Notification, on_delete=models.CASCADE)
    is_read = models.BooleanField(default=False)

    class Meta:
        unique_together = ('user', 'notification')


from django.contrib.auth import get_user_model
User = get_user_model()
class Complaint(models.Model):
    sender = models.ForeignKey(User, on_delete=models.CASCADE, related_name='sent_complaints')
    recipient_type = models.CharField(
        max_length=10,
        choices=[('hr', 'HR'), ('manager', 'Manager')]
    )
    title = models.CharField(max_length=255)
    message = models.TextField()
    
    # ✅ الحقول الخاصة بالرد
    response = models.TextField(blank=True, null=True)
    is_responded = models.BooleanField(default=False)
    responded_at = models.DateTimeField(null=True, blank=True)
    responded_by = models.ForeignKey(  # ✅ من الذي رد
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='responded_complaints'
    )
    is_seen_by_employee = models.BooleanField(default=False)
    is_seen_by_recipient = models.BooleanField(default=False)  # for manager or HR

    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Complaint by {self.sender.username} to {self.recipient_type}"