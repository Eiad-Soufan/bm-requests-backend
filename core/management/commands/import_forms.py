# model_system/management/commands/import_forms.py
import os
from pathlib import Path
from collections import defaultdict

from django.core.management.base import BaseCommand, CommandError
from django.apps import apps
from django.conf import settings
from django.db import transaction
from django.core.files import File

try:
    # نستخدم openpyxl لأنه خفيف ودقيق لقراءة Excel
    from openpyxl import load_workbook
except Exception as e:
    raise CommandError(
        "openpyxl غير مثبت. أضِف 'openpyxl' إلى requirements.txt ثم ثبّته."
    )


def norm(s):
    return (str(s).strip() if s is not None else "").strip()


def normalize_header(h):
    h = norm(h).lower()
    # مرادفات شائعة للأعمدة
    mapping = {
        "serial_number": {"serial", "serial number", "code", "form code", "رقم", "الكود", "serial_number"},
        "name_ar": {"name_ar", "arabic", "arabic name", "الاسم العربي", "name ar"},
        "name_en": {"name_en", "english", "english name", "الاسم الانكليزي", "name en"},
        "category": {"category", "الفئة", "القسم الداخلي", "التصنيف"},
        "description": {"description", "وصف", "details", "ملاحظات"},
        "section": {"section", "القسم", "section_ar", "section_en", "department"},
    }
    for key, aliases in mapping.items():
        if h in aliases:
            return key
    return h  # إبقاء العنوان كما هو إن لم يُعرف


class Command(BaseCommand):
    help = (
        "يمسح مجلد data عن كل ملفات PDF، ثم يقرأ forms.xlsx ويُنشئ/يُحدّث FormModel ويربطه بالقسم المناسب."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            "--data-dir",
            default=str(Path(settings.BASE_DIR) / "data"),
            help="مسار مجلد البيانات الذي يحوي ملفات PDF وملف forms.xlsx (الافتراضي: BASE_DIR/data)",
        )
        parser.add_argument(
            "--excel",
            default="forms.xlsx",
            help="اسم/مسار ملف الإكسل داخل مجلد البيانات (الافتراضي: forms.xlsx)",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="تجربة بدون كتابة في قاعدة البيانات/الملفات",
        )

    def handle(self, *args, **opts):
        data_dir = Path(opts["data_dir"]).resolve()
        excel_path = (data_dir / opts["excel"]).resolve()
        dry_run = opts["dry_run"]

        if not data_dir.exists():
            raise CommandError(f"مجلد البيانات غير موجود: {data_dir}")

        if not excel_path.exists():
            raise CommandError(f"ملف الإكسل غير موجود: {excel_path}")

        # جلب الموديلات بطريقة آمنة بغض النظر عن اسم التطبيق
        Section = apps.get_model("model_system", "Section")
        FormModel = apps.get_model("model_system", "FormModel")

        self.stdout.write(self.style.NOTICE(f"📂 DATA DIR: {data_dir}"))
        self.stdout.write(self.style.NOTICE(f"📄 EXCEL  : {excel_path.name}"))

        # 1) بناء فهرس لملفات PDF: { 'HR-001': '/.../HR-001.pdf', ... }
        pdf_index = {}
        for p in data_dir.glob("*.pdf"):
            code = p.stem.strip()
            pdf_index[code.lower()] = p
        if not pdf_index:
            self.stdout.write(self.style.WARNING("لم يتم العثور على أي PDF في المجلد."))

        # 2) قراءة forms.xlsx إلى قاموس بحسب serial_number
        wb = load_workbook(excel_path, data_only=True)
        # اختيار ورقة اسمها forms إن وجدت، وإلا الورقة النشطة
        ws = wb["forms"] if "forms" in wb.sheetnames else wb.active

        headers = [normalize_header(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))[0:ws.max_column]]
        rows_data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = {headers[i]: row[i] for i in range(min(len(headers), len(row)))}
            # تطبيع
            rd = {
                "serial_number": norm(row_dict.get("serial_number")),
                "name_ar": norm(row_dict.get("name_ar")),
                "name_en": norm(row_dict.get("name_en")),
                "category": norm(row_dict.get("category")),
                "description": norm(row_dict.get("description")),
                "section": norm(row_dict.get("section")),
            }
            if rd["serial_number"]:
                rows_data.append(rd)

        excel_index = {r["serial_number"].lower(): r for r in rows_data}

        created, updated, skipped_no_excel, skipped_no_section, skipped_no_pdf = 0, 0, 0, 0, 0
        problems = defaultdict(list)

        @transaction.atomic
        def do_work():
            nonlocal created, updated, skipped_no_excel, skipped_no_section, skipped_no_pdf

            for code_lower, pdf_path in pdf_index.items():
                # ابحث عن صف الإكسل المطابق
                row = excel_index.get(code_lower)
                if not row:
                    skipped_no_excel += 1
                    problems["لا يوجد صف في الإكسل لهذا الكود"].append(pdf_path.name)
                    continue

                # ابحث عن القسم بالاسم العربي أو الإنكليزي
                section_name = row.get("section")
                section_obj = None
                if section_name:
                    section_obj = (
                        Section.objects.filter(name_ar__iexact=section_name).first()
                        or Section.objects.filter(name_en__iexact=section_name).first()
                    )
                if not section_obj:
                    skipped_no_section += 1
                    problems["القسم غير موجود في قاعدة البيانات"].append(f"{pdf_path.name} -> {section_name!r}")
                    continue

                # تحقق من وجود السجل سابقًا
                obj = FormModel.objects.filter(serial_number__iexact=row["serial_number"]).first()

                if obj:
                    changed = False
                    # تحديث الحقول النصية
                    for fld in ("section", "name_ar", "name_en", "category", "description"):
                        new_val = section_obj if fld == "section" else row.get(fld, "")
                        if getattr(obj, fld) != new_val:
                            setattr(obj, fld, new_val)
                            changed = True

                    # تحديث الملف إن تغيّر الاسم أو الملف غير موجود
                    filename = pdf_path.name
                    if not obj.file or Path(obj.file.name).name != filename:
                        if not dry_run:
                            with open(pdf_path, "rb") as fh:
                                obj.file.save(filename, File(fh), save=False)
                        changed = True

                    if changed and not dry_run:
                        obj.save()
                    updated += 1 if changed else 0
                else:
                    # إنشاء سجل جديد
                    obj = FormModel(
                        section=section_obj,
                        serial_number=row["serial_number"],
                        name_ar=row.get("name_ar", ""),
                        name_en=row.get("name_en", ""),
                        category=row.get("category", ""),
                        description=row.get("description", ""),
                    )
                    if not dry_run:
                        with open(pdf_path, "rb") as fh:
                            obj.file.save(pdf_path.name, File(fh), save=False)
                        obj.save()
                    created += 1

            # تحذير للأسطر في الإكسل التي لا يقابلها ملف PDF
            for code_lower in set(excel_index.keys()) - set(pdf_index.keys()):
                skipped_no_pdf += 1
                problems["لا يوجد PDF لهذا الكود"].append(excel_index[code_lower]["serial_number"])

        if dry_run:
            self.stdout.write(self.style.WARNING("تشغيل تجريبي DRY RUN — لن يتم أي حفظ."))
        do_work()

        # ملخص
        self.stdout.write("")
        self.stdout.write(self.style.SUCCESS(f"✅ Created: {created}"))
        self.stdout.write(self.style.SUCCESS(f"✅ Updated: {updated}"))
        self.stdout.write(self.style.WARNING(f"⛔ Skipped (no excel row): {skipped_no_excel}"))
        self.stdout.write(self.style.WARNING(f"⛔ Skipped (section missing): {skipped_no_section}"))
        self.stdout.write(self.style.WARNING(f"⛔ Skipped (no PDF for excel row): {skipped_no_pdf}"))

        if problems:
            self.stdout.write("\nتفاصيل المشاكل:")
            for k, items in problems.items():
                for it in items[:20]:  # لا نغرق الشاشة
                    self.stdout.write(f" - {k}: {it}")
            if any(len(v) > 20 for v in problems.values()):
                self.stdout.write("... (تم تقصير القائمة)")


