from rest_framework import viewsets, status
from rest_framework.response import Response
from rest_framework.decorators import action, api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.views import APIView
from django.shortcuts import get_object_or_404
from django.http import FileResponse, Http404
from django.utils import timezone

from .models import Notification, UserNotification, Section, FormModel, Complaint
from .serializers import (
    SectionSerializer,
    FormModelSerializer,
    NotificationSerializer,
    UserNotificationSerializer,
    ComplaintSerializer,
    MyTokenObtainPairSerializer
)

from rest_framework_simplejwt.views import TokenObtainPairView
from django.contrib.auth import get_user_model
User = get_user_model()

# 🔑 توكين JWT مخصص لإرجاع صلاحيات المستخدم
class MyTokenObtainPairView(TokenObtainPairView):
    serializer_class = MyTokenObtainPairSerializer


# 🔐 معلومات المستخدم الحالي
@api_view(['GET'])
@permission_classes([IsAuthenticated])
def current_user_info(request):
    user = request.user
    return Response({
        'username': user.username,
        'is_staff': user.is_staff,
        'is_superuser': user.is_superuser,
        'id': user.id,
        'email': user.email,
        'role':  user.role
    })


# 📋 عرض المستخدمين بالأسماء لاختيار الإشعار
class UserListAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request):
        users = User.objects.all().values('id', 'username', 'email')
        return Response(list(users))


# 🌐 عرض النموذج للعامة بدون حماية
@api_view(['GET'])
def public_form_preview(request, pk):
    try:
        form = FormModel.objects.get(pk=pk)
        return FileResponse(form.file.open(), content_type='application/pdf')
    except FormModel.DoesNotExist:
        raise Http404("Form not found")


# 📄 عرض النموذج الداخلي
def preview_form(request, form_id):
    form = get_object_or_404(FormModel, id=form_id)
    response = FileResponse(form.file.open('rb'), content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="form.pdf"'
    return response


# 🔔 إرسال إشعار لمستخدمين أو للجميع
class NotificationViewSet(viewsets.ModelViewSet):
    queryset = Notification.objects.all()
    serializer_class = NotificationSerializer
    permission_classes = [IsAuthenticated]

    @action(detail=False, methods=['post'])
    def send_notification(self, request):
        print(request.data)
        title = request.data.get('title')
        message = request.data.get('message')
        importance = request.data.get('importance')
        usernames = request.data.get('usernames')  # قائمة الأسماء

        notification = Notification.objects.create(
            title=title,
            message=message,
            importance=importance
        )

        if usernames:
            users = User.objects.filter(username__in=usernames)
        else:
            users = User.objects.all()

        UserNotification.objects.bulk_create([
            UserNotification(user=user, notification=notification) for user in users
        ])

        return Response({'status': 'Notification sent successfully'}, status=status.HTTP_201_CREATED)


# 📂 عرض الأقسام (Tabs)
class SectionViewSet(viewsets.ReadOnlyModelViewSet):
    queryset = Section.objects.all()
    serializer_class = SectionSerializer
    permission_classes = [IsAuthenticated]


# 🗂️ عرض النماذج داخل كل قسم
class FormModelViewSet(viewsets.ReadOnlyModelViewSet):
    queryset = FormModel.objects.all()
    serializer_class = FormModelSerializer
    permission_classes = [IsAuthenticated]

    def get_queryset(self):
        user = self.request.user
        # المدير والموارد البشرية يمكنهم الوصول لكل النماذج
        if hasattr(user, 'profile') and user.profile.role in ['manager', 'hr']:
            return FormModel.objects.all()
        allowed_sections = user.usersectionpermission_set.values_list('section_id', flat=True)
        return FormModel.objects.filter(section__id__in=allowed_sections)


# 📩 إشعارات المستخدم الفردية
class UserNotificationViewSet(viewsets.ViewSet):
    permission_classes = [IsAuthenticated]

    def list(self, request):
        user_notifications = UserNotification.objects.filter(
            user=request.user
        ).order_by('-notification__created_at')

        serializer = UserNotificationSerializer(user_notifications, many=True)
        return Response(serializer.data)

    @action(detail=True, methods=['post'])
    def mark_as_read(self, request, pk=None):
        try:
            user_notification = UserNotification.objects.get(pk=pk, user=request.user)
            user_notification.is_read = True
            user_notification.save()
            return Response({'status': 'Marked as read'})
        except UserNotification.DoesNotExist:
            return Response({'error': 'Not found'}, status=status.HTTP_404_NOT_FOUND)

# 📝 API مخصصة للشكاوى
# ====== داخل core/views.py: استبدل كتلة ComplaintViewSet بالكامل بما يلي ======
class ComplaintViewSet(viewsets.ViewSet):
    permission_classes = [IsAuthenticated]

    # 1) إرسال شكوى من موظف
    @action(detail=False, methods=['post'])
    def submit(self, request):
        serializer = ComplaintSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        # الموظف رأى شكواه لحظة الإرسال، والجهة المستقبلة تراها غير مقروءة
        complaint = serializer.save(
            sender=request.user,
            is_responded=False,
            is_seen_by_recipient=False,
            is_seen_by_employee=True
        )
        return Response(ComplaintSerializer(complaint).data, status=status.HTTP_201_CREATED)

    # 2) شكاوى الموظف الحالي
    @action(detail=False, methods=['get'])
    def my_complaints(self, request):
        qs = Complaint.objects.filter(sender=request.user).order_by('-created_at')
        return Response(ComplaintSerializer(qs, many=True).data)

    # 3) شكاوى موجّهة للـ HR
    @action(detail=False, methods=['get'])
    def hr_complaints(self, request):
        qs = Complaint.objects.filter(recipient_type='hr').order_by('-created_at')
        return Response(ComplaintSerializer(qs, many=True).data)

    # 4) شكاوى موجّهة للمدير
    @action(detail=False, methods=['get'])
    def manager_complaints(self, request):
        qs = Complaint.objects.filter(recipient_type='manager').order_by('-created_at')
        return Response(ComplaintSerializer(qs, many=True).data)

    # 5) رد HR على شكوى
    @action(detail=True, methods=['post'])
    def hr_reply(self, request, pk=None):
        complaint = get_object_or_404(Complaint, pk=pk, recipient_type='hr')
        response_text = request.data.get('response')
        if not response_text:
            return Response({'error': 'Response is required'}, status=400)

        complaint.response = response_text
        complaint.is_responded = True
        complaint.responded_by = request.user
        complaint.responded_at = timezone.now()
        complaint.is_seen_by_recipient = True     # الجهة المعالجة قرأتها
        complaint.is_seen_by_employee = False     # الموظف لديه رد جديد غير مقروء
        complaint.save(update_fields=[
            'response','is_responded','responded_by','responded_at',
            'is_seen_by_recipient','is_seen_by_employee'
        ])
        return Response({'status': 'Response saved'})

    # 6) رد المدير على شكوى
    @action(detail=True, methods=['post'])
    def manager_reply(self, request, pk=None):
        complaint = get_object_or_404(Complaint, pk=pk, recipient_type='manager')
        response_text = request.data.get('response')
        if not response_text:
            return Response({'error': 'Response is required'}, status=400)

        complaint.response = response_text
        complaint.is_responded = True
        complaint.responded_by = request.user
        complaint.responded_at = timezone.now()
        complaint.is_seen_by_recipient = True
        complaint.is_seen_by_employee = False
        complaint.save(update_fields=[
            'response','is_responded','responded_by','responded_at',
            'is_seen_by_recipient','is_seen_by_employee'
        ])
        return Response({'status': 'Response saved'})

    # 7) تعليم شكوى واحدة كمقروءة حسب الدور
    @action(detail=True, methods=['post'])
    def mark_seen(self, request, pk=None):
        complaint = get_object_or_404(Complaint, pk=pk)
        user = request.user
        role = getattr(user, 'role', None)

        if user == complaint.sender:
            complaint.is_seen_by_employee = True
            fields = ['is_seen_by_employee']
        elif role in ['manager', 'hr'] and complaint.recipient_type == role:
            complaint.is_seen_by_recipient = True
            fields = ['is_seen_by_recipient']
        else:
            return Response({'error': 'Not allowed'}, status=403)

        complaint.save(update_fields=fields)
        return Response({'message': 'Marked as seen'})

    # 8) تعليم الكل كمقروء (مسار يطلبه الفرونت: /api/complaints/mark_all_seen/)
    @action(detail=False, methods=['post'])
    def mark_all_seen(self, request):
        user = request.user
        role = getattr(user, 'role', None)

        if role in ['manager', 'hr']:
            Complaint.objects.filter(
                recipient_type=role,
                is_seen_by_recipient=False
            ).update(is_seen_by_recipient=True)
        else:
            Complaint.objects.filter(
                sender=user,
                is_responded=True,
                is_seen_by_employee=False
            ).update(is_seen_by_employee=True)

        return Response({'message': 'OK'})


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def has_unread_complaints(request):
    """
    المدير/HR: أي شكاوى موجّهة إليهم ولم تُقرأ بعد.
    الموظف: فقط الشكاوى التي تم الرد عليها ولم يقرأها الموظف بعد.
    """
    user = request.user
    role = getattr(user, 'role', None)

    if role == 'manager':
        has_new = Complaint.objects.filter(
            recipient_type='manager',
            is_seen_by_recipient=False
        ).exists()
    elif role == 'hr':
        has_new = Complaint.objects.filter(
            recipient_type='hr',
            is_seen_by_recipient=False
        ).exists()
    else:
        has_new = Complaint.objects.filter(
            sender=user,
            is_responded=True,
            is_seen_by_employee=False
        ).exists()

    return Response({'has_new': has_new})


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def mark_complaint_as_seen(request, pk):
    """
    بديل/مرادف للـ action أعلاه إذا أردت إبقاء هذا المسار القديم يعمل أيضًا:
    /api/complaints/<pk>/mark_seen/
    """
    complaint = get_object_or_404(Complaint, pk=pk)
    user = request.user
    role = getattr(user, 'role', None)

    if user == complaint.sender:
        complaint.is_seen_by_employee = True
        fields = ['is_seen_by_employee']
    elif role in ['manager', 'hr'] and complaint.recipient_type == role:
        complaint.is_seen_by_recipient = True
        fields = ['is_seen_by_recipient']
    else:
        return Response({'error': 'Not allowed'}, status=403)

    complaint.save(update_fields=fields)
    return Response({'message': 'Marked as seen'})


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def mark_all_complaints_seen(request):
    """
    مسار علوي قديم (موجود في urls.py باسم mark-all-complaints-seen/).
    أبقيناه لكنه الآن يحدّث الحقول الصحيحة.
    """
    user = request.user
    role = getattr(user, 'role', None)

    if role == 'manager':
        Complaint.objects.filter(
            recipient_type='manager',
            is_seen_by_recipient=False
        ).update(is_seen_by_recipient=True)
    elif role == 'hr':
        Complaint.objects.filter(
            recipient_type='hr',
            is_seen_by_recipient=False
        ).update(is_seen_by_recipient=True)
    else:
        Complaint.objects.filter(
            sender=user,
            is_responded=True,
            is_seen_by_employee=False
        ).update(is_seen_by_employee=True)

    return Response({'status': 'All marked as seen'})