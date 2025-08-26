from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import (
    SectionViewSet,
    FormModelViewSet,
    NotificationViewSet,
    UserNotificationViewSet,
    MyTokenObtainPairView,
    UserListAPIView,
    current_user_info,
    preview_form,
    public_form_preview,
    ComplaintViewSet,
)
from rest_framework_simplejwt.views import TokenRefreshView
from .views import *

user_notifications = UserNotificationViewSet.as_view({
    'get': 'list',
})

mark_as_read = UserNotificationViewSet.as_view({
    'post': 'mark_as_read',
})

send_notification = NotificationViewSet.as_view({
    'post': 'send_notification',
})

router = DefaultRouter()
router.register(r'sections', SectionViewSet, basename='section')
router.register(r'forms', FormModelViewSet, basename='formmodel')
router.register(r'notifications', NotificationViewSet, basename='notification')
router.register(r'user-notifications', UserNotificationViewSet, basename='user-notifications')
router.register(r'complaints', ComplaintViewSet, basename='complaint')

urlpatterns = [
    path('', include(router.urls)),
    path('token/', MyTokenObtainPairView.as_view(), name='token_obtain_pair'),
    path('token/refresh/', TokenRefreshView.as_view(), name='token_refresh'),
    path('users/', UserListAPIView.as_view(), name='user-list'),
    path('me/', current_user_info, name='current-user-info'),
    path('preview-form/<int:form_id>/', preview_form, name='preview-form'),
    path('public-form/<int:pk>/', public_form_preview, name='public-form-preview'),
    path('current-user/', current_user_info, name='current-user'),
    path('complaints/<int:pk>/mark_seen/', mark_complaint_as_seen),
    path('complaints/has_unread/', has_unread_complaints, name='has-unread-complaints'),
    path('mark-all-complaints-seen/', mark_all_complaints_seen, name='mark_all_complaints_seen'),


]
urlpatterns += [
    path('user-notifications/', user_notifications, name='user-notifications-list'),
    path('user-notifications/<int:pk>/mark_as_read/', mark_as_read, name='user-notifications-mark-as-read'),
    path('notify-admin/send_notification/', send_notification, name='send_notification'),
]