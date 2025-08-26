from rest_framework import serializers
from .models import Section, FormModel
from .models import Notification, UserNotification
from rest_framework_simplejwt.serializers import TokenObtainPairSerializer
from .models import Complaint

class MyTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        token = super().get_token(user)

        # نضيف معلومات إضافية داخل التوكن
        token['username'] = user.username
        token['is_staff'] = user.is_staff
        token['is_superuser'] = user.is_superuser

        return token
    
class SectionSerializer(serializers.ModelSerializer):
    class Meta:
        model = Section
        fields = ['id', 'name_ar', 'name_en']

class FormModelSerializer(serializers.ModelSerializer):
    section = SectionSerializer(read_only=True)

    class Meta:
        model = FormModel
        fields = [
            'id', 'serial_number', 'name_ar', 'name_en',
            'category', 'description', 'file', 'section'
        ]


class NotificationSerializer(serializers.ModelSerializer):
    importance_display = serializers.CharField(source='get_importance_display', read_only=True)

    class Meta:
        model = Notification
        fields = ['id', 'title', 'message', 'importance', 'importance_display', 'created_at']

class UserNotificationSerializer(serializers.ModelSerializer):
    notification = NotificationSerializer()

    class Meta:
        model = UserNotification
        fields = ['id', 'notification', 'is_read']

class ComplaintSerializer(serializers.ModelSerializer):
    sender_username = serializers.CharField(source='sender.username', read_only=True)
    recipient_display = serializers.SerializerMethodField()
    is_seen_by_employee = serializers.BooleanField(read_only=True)
    is_seen_by_recipient = serializers.BooleanField(read_only=True)

    class Meta:
        model = Complaint
        fields = '__all__'
        read_only_fields = [
            'sender',
            'sender_username',
            'created_at',
            'is_responded',
            'response',
            'responded_at',
            'recipient_display'
        ]

    def get_recipient_display(self, obj):
        return obj.get_recipient_type_display()