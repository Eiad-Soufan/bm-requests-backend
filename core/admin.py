from django.contrib import admin
from .models import Section, FormModel, UserSectionPermission, Notification, UserNotification
from django.contrib.auth import get_user_model
from .models import Complaint




CustomUser = get_user_model()

@admin.register(Section)
class SectionAdmin(admin.ModelAdmin):
    list_display = ('id', 'name_ar', 'name_en')

@admin.register(FormModel)
class FormModelAdmin(admin.ModelAdmin):
    list_display = ('serial_number', 'name_ar', 'section', 'category')
    list_filter = ('section', 'category')
    search_fields = ('name_ar', 'name_en', 'serial_number')

@admin.register(UserSectionPermission)
class UserSectionPermissionAdmin(admin.ModelAdmin):
    list_display = ('user', 'section')

@admin.register(Notification)
class NotificationAdmin(admin.ModelAdmin):
    list_display = ('title', 'message', 'importance', 'created_at')

@admin.register(UserNotification)
class UserNotificationAdmin(admin.ModelAdmin):
    list_display = ('user', 'notification')

@admin.register(CustomUser)
class CustomUserAdmin(admin.ModelAdmin):
    list_display = ('username', 'email', 'role', 'is_staff', 'is_superuser')
    list_filter = ('role', 'is_staff')
    search_fields = ('username', 'email')
    
@admin.register(Complaint)
class ComplaintAdmin(admin.ModelAdmin):
    list_display = ('title','message','response', 'sender', 'recipient_type', 'is_responded', 'created_at')