# management/commands/import_forms.py
import re
from pathlib import Path
from collections import defaultdict

from django.core.management.base import BaseCommand, CommandError
from django.apps import apps
from django.conf import settings
from django.db import transaction
from django.core.files import File

try:
    from openpyxl import load_workbook
except Exception:
    raise CommandError("openpyxl غير مثبت. أضِفه إلى requirements.txt ثم ثبّته.")

# تصحيح أسماء الشيتات الشائعة/الأخطاء الإملائية
SHEET_ALIASES = {
    "human recourses": "Human Resources",
    "purchasing management": "Purchasing Management",
    "maintenance department": "Maintenance Department",
    "public relations": "Public Relations",
    "production department": "Production Department",
    "marketing department": "Marketing Department",
    "it and development": "IT and Development",
    "wholesale management": "Wholesale Management",
    "financial management": "Financial Management",
    "retail management": "Retail Management",
}

# خريطة بادئة الكود → القسم الافتراضي (عند غياب صف في الإكسل)
CODE_PREFIX_TO_SECTION_EN = {
    "HR": "Human Resources",
    "FN": "Financial Management",
    "WS": "Wholesale Management",
    "RT": "Retail Management",
    "PR": "Purchasing Management",
    "NT": "Maintenance Department",
    "PD": "Production Department",
    "MK": "Marketing Department",
    "IT": "IT and Development",
    "AG": "Agents Service",
}

def norm(s):
    return (str(s).strip() if s is not None else "").strip()

def normalize_header(h):
    h = norm(h).lower()
    mapping = {
        "serial_number": {"serial", "serial number", "form no", "form number", "form id",
                          "code", "form code", "رقم", "الكود", "serial_number", "id"},
        "name_ar": {"name_ar", "arabic", "arabic name", "الاسم العربي", "name ar", "name (arabic)"},
        "name_en": {"name_en", "english", "english name", "الاسم الانكليزي", "name en", "name (english)"},
        "category": {"category", "الفئة", "القسم الداخلي", "التصنيف"},
        "description": {"description", "وصف", "details", "ملاحظات", "desc"},
        "section": {"section", "القسم", "section_ar", "section_en", "department"},
    }
    for key, aliases in mapping.items():
        if h in aliases: return key
    return h

DASHES = {"\u2013", "\u2014", "_"}  # – — _
def norm_code(code: str) -> str:
    s = norm(code)
    if not s: return ""
    s = s.replace(".PDF", "").replace(".pdf", "")
    for d in DASHES: s = s.replace(d, "-")
    s = re.sub(r"\s+", "", s).lower()
    m = re.match(r"^([a-z]+)-?(\d+)$", s)
    return f"{m.group(1)}-{m.group(2).zfill(3)}" if m else s

def sheet_clean_name(name: str) -> str:
    base = re.sub(r"\s+", " ", norm(name).strip().strip("-"))
    return SHEET_ALIASES.get(base.lower()) or base

def has_fields(model, needed):
    return needed.issubset({f.name for f in model._meta.get_fields() if hasattr(f, "name")})

def find_models(app_label=None):
    if app_label:
        try:
            return apps.get_model(app_label, "Section"), apps.get_model(app_label, "FormModel")
        except LookupError as e:
            raise CommandError(f"لا يوجد app '{app_label}' يحوي Section/FormModel. ({e})")
    Section = FormModel = None
    for m in apps.get_models():
        if m.__name__ == "Section" and has_fields(m, {"name_ar","name_en"}): Section = m
        if m.__name__ == "FormModel" and has_fields(m, {"serial_number","name_ar","name_en","category","description","file"}): FormModel = m
    if not Section or not FormModel:
        labels = [a.label for a in apps.get_app_configs()]
        raise CommandError("تعذّر إيجاد Section/FormModel تلقائيًا. مرّر --app-label.\n"
                           f"التطبيقات المتاحة: {', '.join(labels)}")
    return Section, FormModel

def detect_header_row(ws, max_scan=20):
    wanted = {"serial_number","name_ar","name_en","category","description","section"}
    best = (0, 1, [])
    for r in range(1, min(ws.max_row, max_scan)+1):
        cells = [normalize_header(c.value) for c in ws[r]]
        score = sum(1 for v in cells if v in wanted) + (2 if "serial_number" in cells else 0)
        if score > best[0]: best = (score, r, cells)
    return best

def guess_section_from_code(code: str):
    m = re.match(r"^([A-Za-z]+)", norm(code))
    return CODE_PREFIX_TO_SECTION_EN.get(m.group(1).upper()) if m else None

def find_or_create_section(Section, name, create_missing=True):
    if not name: return None
    name = sheet_clean_name(name)
    obj = (Section.objects.filter(name_ar__iexact=name).first()
           or Section.objects.filter(name_en__iexact=name).first()
           or Section.objects.filter(name_en__icontains=name).first()
           or Section.objects.filter(name_ar__icontains=name).first())
    return obj or (Section.objects.create(name_ar=name, name_en=name) if create_missing else None)

class Command(BaseCommand):
    help = "يستورد ملفات PDF من data/ ويربطها بصفوف forms.xlsx (كل الشيتات) ويُنشئ الأقسام الناقصة، ولن يترك أي PDF دون إدخال."

    def add_arguments(self, parser):
        parser.add_argument("--data-dir", default=str(Path(settings.BASE_DIR)/"data"),
                            help="مجلد البيانات (افتراضي BASE_DIR/data)")
        parser.add_argument("--excel", default="forms.xlsx",
                            help="اسم/مسار ملف الإكسل داخل مجلد البيانات")
        parser.add_argument("--sheet", help="قراءة شيت واحد فقط (اختياري)")
        parser.add_argument("--dry-run", action="store_true", help="تشغيل تجريبي بلا حفظ")
        parser.add_argument("--app-label", help="وسم التطبيق الذي يحوي Section و FormModel (مثال: core)")
        parser.add_argument("--no-create-sections", action="store_true", help="عدم إنشاء أقسام جديدة تلقائيًا.")

    def handle(self, *args, **opts):
        data_dir = Path(opts["data_dir"]).resolve()
        excel_path = (data_dir / opts["excel"]).resolve()
        sheet_only = opts.get("sheet")
        dry_run = opts["dry_run"]
        app_label = opts.get("app_label")
        create_missing_sections = not opts["no_create_sectors"] if "no_create_sectors" in opts else not opts["no_create_sections"]

        if not data_dir.exists(): raise CommandError(f"مجلد البيانات غير موجود: {data_dir}")
        if not excel_path.exists(): raise CommandError(f"ملف الإكسل غير موجود: {excel_path}")

        Section, FormModel = find_models(app_label)

        self.stdout.write(self.style.NOTICE(f"📂 DATA DIR: {data_dir}"))
        self.stdout.write(self.style.NOTICE(f"📄 EXCEL  : {excel_path.name}"))

        # فهرسة ملفات PDF
        pdf_index = {}
        for p in sorted(data_dir.glob("*.pdf")):
            key = norm_code(p.stem)
            if key: pdf_index[key] = p
        if not pdf_index:
            self.stdout.write(self.style.WARNING("لم يتم العثور على أي PDF في المجلد."))

        # قراءة الإكسل
        wb = load_workbook(excel_path, data_only=True)
        sheetnames = [sheet_only] if sheet_only else wb.sheetnames

        rows_data = []
        for sname in sheetnames:
            ws = wb[sname]
            score, header_row_idx, headers = detect_header_row(ws)
            if score == 0:
                self.stdout.write(self.style.WARNING(f"تخطي '{sname}' لعدم العثور على صف عناوين مناسب."))
                continue

            clean_sheet_name = sheet_clean_name(sname)
            for row in ws.iter_rows(min_row=header_row_idx+1, values_only=True):
                if not row: continue
                row_dict = {}
                for i in range(len(headers)):
                    key = normalize_header(headers[i]) if i < len(headers) else None
                    if key: row_dict[key] = row[i] if i < len(row) else None

                serial = norm(row_dict.get("serial_number"))
                if not serial: continue

                rows_data.append({
                    "serial_number": serial,
                    "serial_key": norm_code(serial),
                    "name_ar": norm(row_dict.get("name_ar")),
                    "name_en": norm(row_dict.get("name_en")),
                    "category": norm(row_dict.get("category")),
                    "description": norm(row_dict.get("description")),
                    "section": norm(row_dict.get("section")) or clean_sheet_name,
                })

        excel_index = {r["serial_key"]: r for r in rows_data if r["serial_key"]}
        self.stdout.write(self.style.NOTICE(f"🧾 Rows loaded: {len(rows_data)} from {len(sheetnames)} sheet(s)."))
        self.stdout.write(self.style.NOTICE(f"📑 PDFs found: {len(pdf_index)}"))

        created = updated = used_fallback = 0
        skipped_bad_code = 0
        problems = defaultdict(list)

        @transaction.atomic
        def do_work():
            nonlocal created, updated, used_fallback, skipped_bad_code
            for key, pdf_path in pdf_index.items():
                if not key:
                    skipped_bad_code += 1
                    problems["اسم ملف غير صالح"].append(pdf_path.name)
                    continue

                row = excel_index.get(key)
                if row:
                    section_name = row["section"]
                else:
                    # لا يوجد صف في الإكسل → أنشئ سجلًا اعتمادًا على اسم الملف/بادئة الكود
                    section_name = guess_section_from_code(pdf_path.stem) or "Uncategorized"
                    used_fallback += 1
                    row = {
                        "serial_number": pdf_path.stem,
                        "serial_key": key,
                        "name_ar": "",
                        "name_en": "",
                        "category": "",
                        "description": "",
                        "section": section_name,
                    }

                section_obj = find_or_create_section(Section, section_name, create_missing=create_missing_sections)
                if not section_obj:
                    problems["تعذر تحديد/إنشاء قسم"].append(f"{pdf_path.name} -> {section_name!r}")
                    continue

                obj = FormModel.objects.filter(serial_number__iexact=row["serial_number"]).first()
                if obj:
                    changed = False
                    if obj.section_id != section_obj.id:
                        obj.section = section_obj; changed = True
                    for fld in ("name_ar","name_en","category","description"):
                        val = row.get(fld, "")
                        if getattr(obj, fld) != val:
                            setattr(obj, fld, val); changed = True
                    filename = pdf_path.name
                    if not obj.file or Path(obj.file.name).name != filename:
                        if not dry_run:
                            with open(pdf_path, "rb") as fh:
                                obj.file.save(filename, File(fh), save=False)
                        changed = True
                    if changed and not dry_run:
                        obj.save()
                    if changed: updated += 1
                else:
                    obj = FormModel(
                        section=section_obj,
                        serial_number=row["serial_number"],
                        name_ar=row.get("name_ar",""),
                        name_en=row.get("name_en",""),
                        category=row.get("category",""),
                        description=row.get("description",""),
                    )
                    if not dry_run:
                        with open(pdf_path, "rb") as fh:
                            obj.file.save(pdf_path.name, File(fh), save=False)
                        obj.save()
                    created += 1

        if dry_run:
            self.stdout.write(self.style.Warning("DRY RUN — لن يتم أي حفظ."))
        do_work()

        self.stdout.write("")
        self.stdout.write(self.style.SUCCESS(f"✅ Created: {created}"))
        self.stdout.write(self.style.SUCCESS(f"✅ Updated: {updated}"))
        self.stdout.write(self.style.SUCCESS(f"🤝 Used fallback (no Excel row): {used_fallback}"))
        self.stdout.write(self.style.WARNING(f"⛔ Skipped (bad code): {skipped_bad_code}"))

        if problems:
            self.stdout.write("\nتفاصيل إضافية:")
            for k, items in problems.items():
                for it in items[:80]:
                    self.stdout.write(f" - {k}: {it}")
            if any(len(v) > 80 for v in problems.values()):
                self.stdout.write("... (تم تقصير القائمة)")
