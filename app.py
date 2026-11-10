
import streamlit as st
import pandas as pd
import re
import io
import json
import unicodedata
from collections import Counter
from rapidfuzz import fuzz, process
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Smart Excel Matcher",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
html, body, [class*="css"] {
    direction: rtl;
    text-align: right;
}
.stApp { direction: rtl; }
[data-testid="stDataFrame"] { direction: rtl; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Smart Excel Matcher")
st.caption("نظام متقدم لمطابقة ومقارنة ملفات Excel")

WEIGHTS = {
    "الاسم": 40,
    "نوع الهوية": 10,
    "رقم الهوية": 30,
    "رقم الحساب": 20
}


# -------------------------
# أدوات تنظيف البيانات
# -------------------------

def clean_text(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    value = unicodedata.normalize("NFKC", value)
    value = re.sub(r"[\u064B-\u065F\u0670]", "", value)
    value = value.replace("ـ", "")
    value = value.translate(str.maketrans(
        "أإآىؤئ", "ااايوي"
    ))
    value = re.sub(r"\s+", " ", value)
    return value.lower().strip()


def clean_number(value):
    value = clean_text(value)
    arabic = "٠١٢٣٤٥٦٧٨٩"
    persian = "۰۱۲۳۴۵۶۷۸۹"
    english = "0123456789"
    value = value.translate(str.maketrans(arabic, english))
    value = value.translate(str.maketrans(persian, english))
    value = re.sub(r"\s+", "", value)
    return value


def normalize_value(value, numeric=False):
    return clean_number(value) if numeric else clean_text(value)


def read_uploaded_excel(uploaded_file, label):
    excel = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox(
        f"ورقة العمل - {label}",
        excel.sheet_names,
        key=f"sheet_{label}"
    )
    df = pd.read_excel(
        uploaded_file,
        sheet_name=sheet,
        dtype=str
    ).fillna("")
    return df.reset_index(drop=True)


def build_key(row, columns):
    values = [clean_number(row[c]) for c in columns]
    if not values or any(v == "" for v in values):
        return ""
    return "||".join(values)


def find_duplicates(df, columns):
    if not columns:
        return pd.DataFrame()
    keys = df.apply(lambda r: build_key(r, columns), axis=1)
    mask = (keys != "") & keys.duplicated(keep=False)
    result = df.loc[mask].copy()
    if not result.empty:
        result.insert(0, "رقم الصف في الملف", result.index + 2)
        result.insert(1, "مفتاح التكرار", keys.loc[mask])
    return result.reset_index(drop=True)


def weighted_score(row_a, row_b, fields):
    total = 0
    used_weight = 0
    details = {}

    for field, (col_a, col_b) in fields.items():
        a = row_a[col_a]
        b = row_b[col_b]
        numeric = field in ["رقم الهوية", "رقم الحساب"]

        a = normalize_value(a, numeric)
        b = normalize_value(b, numeric)

        if not a or not b:
            details[field] = "غير متوفر"
            continue

        score = fuzz.ratio(a, b)
        weight = WEIGHTS[field]
        total += score * weight
        used_weight += weight
        details[field] = round(score, 1)

    score = total / used_weight if used_weight else 0
    return round(score, 1), details


def score_label(score):
    if score >= 95:
        return "مرتفع جداً"
    if score >= 85:
        return "مرتفع"
    if score >= 70:
        return "يحتاج مراجعة"
    return "ضعيف"


def compare_fields(row_a, row_b, fields):
    equal_fields = []
    different_fields = []

    for field, (col_a, col_b) in fields.items():
        numeric = field in ["رقم الهوية", "رقم الحساب"]
        a = normalize_value(row_a[col_a], numeric)
        b = normalize_value(row_b[col_b], numeric)

        if not a or not b:
            continue

        if a == b:
            equal_fields.append(field)
        else:
            different_fields.append(field)

    return "، ".join(equal_fields), "، ".join(different_fields)


# -------------------------
# المطابقة الدقيقة
# -------------------------

def exact_match(source, target, source_cols, target_cols,
                target_output_cols, direction):

    target_keys = target.apply(
        lambda r: build_key(r, target_cols), axis=1
    )
    source_keys = source.apply(
        lambda r: build_key(r, source_cols), axis=1
    )

    lookup = {}
    for idx, key in target_keys.items():
        if key:
            lookup.setdefault(key, []).append(idx)

    records = []
    matched_target = set()

    for idx, row in source.iterrows():
        key = source_keys.loc[idx]
        matches = lookup.get(key, [])

        record = {
            "اتجاه المطابقة": direction,
            "صف الملف المرجعي": idx + 2,
            "مفتاح المطابقة": key
        }

        for col in source.columns:
            record[f"الملف المرجعي - {col}"] = row[col]

        if not key:
            status = "مفتاح فارغ"
            target_idx = None
        elif len(matches) > 1:
            status = "مفتاح مكرر في الملف المقابل"
            target_idx = None
        elif len(matches) == 1:
            status = "متطابق"
            target_idx = matches[0]
            matched_target.add(target_idx)
        else:
            status = "غير موجود في الملف المقابل"
            target_idx = None

        record["حالة المطابقة"] = status
        record["صف الملف المقابل"] = (
            target_idx + 2 if target_idx is not None else ""
        )

        for col in target_output_cols:
            record[f"الملف المقابل - {col}"] = (
                target.loc[target_idx, col]
                if target_idx is not None else ""
            )

        records.append(record)

    results = pd.DataFrame(records)

    unmatched = target.loc[
        ~target.index.isin(matched_target)
    ].copy()

    if not unmatched.empty:
        unmatched.insert(0, "صف الملف المقابل", unmatched.index + 2)
        unmatched.insert(1, "اتجاه المطابقة", direction)

    return results, unmatched


# -------------------------
# المطابقة الذكية
# -------------------------

def smart_match(source, target, fields, output_cols,
                threshold, one_to_one, direction):

    target_names = []
    name_field = fields.get("الاسم")

    if name_field:
        target_names = [
            clean_text(v) for v in target[name_field[1]].tolist()
        ]

    used_target = set()
    matched_target = set()
    records = []

    for idx, row in source.iterrows():
        candidates = set()

        # مرشحون من أرقام الهوية والحساب
        for field in ["رقم الهوية", "رقم الحساب"]:
            if field in fields:
                source_col, target_col = fields[field]
                value = clean_number(row[source_col])
                if value:
                    matches = target.index[
                        target[target_col].apply(clean_number) == value
                    ].tolist()
                    candidates.update(matches)

        # مرشحون من الاسم
        if name_field:
            query = clean_text(row[name_field[0]])
            if query and target_names:
                found = process.extract(
                    query,
                    target_names,
                    scorer=fuzz.ratio,
                    limit=15
                )
                for _, score, pos in found:
                    if score >= 40:
                        candidates.add(target.index[pos])

        # عند غياب المرشحين، نفحص الملف كاملاً إذا كان حجمه مناسباً
        if not candidates and len(target) <= 3000:
            candidates = set(target.index.tolist())

        best_idx = None
        best_score = -1
        best_details = {}

        for candidate_idx in candidates:
            if one_to_one and candidate_idx in used_target:
                continue

            score, details = weighted_score(
                row, target.loc[candidate_idx], fields
            )

            if score > best_score:
                best_score = score
                best_idx = candidate_idx
                best_details = details

        record = {
            "اتجاه المطابقة": direction,
            "صف الملف المرجعي": idx + 2
        }

        for col in source.columns:
            record[f"الملف المرجعي - {col}"] = row[col]

        if best_idx is None:
            status = "غير موجود"
            record["نسبة التطابق"] = 0
            record["تصنيف التطابق"] = "غير موجود"
            record["الحقول المتطابقة"] = ""
            record["الحقول المختلفة"] = ""
            record["صف الملف المقابل"] = ""
        else:
            target_row = target.loc[best_idx]
            equal, different = compare_fields(row, target_row, fields)

            record["نسبة التطابق"] = best_score
            record["تصنيف التطابق"] = score_label(best_score)
            record["الحقول المتطابقة"] = equal
            record["الحقول المختلفة"] = different
            record["صف الملف المقابل"] = best_idx + 2

            if best_score < threshold:
                status = "يحتاج مراجعة"
            elif best_score >= 95:
                status = "تطابق مرتفع جداً"
                used_target.add(best_idx)
                matched_target.add(best_idx)
            else:
                status = "تطابق مرتفع - مراجعة"
                used_target.add(best_idx)
                matched_target.add(best_idx)

        record["حالة المطابقة"] = status

        for field in WEIGHTS:
            record[f"نسبة {field}"] = best_details.get(field, "")

        for col in output_cols:
            record[f"الملف المقابل - {col}"] = (
                target.loc[best_idx, col]
                if best_idx is not None else ""
            )

        records.append(record)

    results = pd.DataFrame(records)

    unmatched = target.loc[
        ~target.index.isin(matched_target)
    ].copy()

    if not unmatched.empty:
        unmatched.insert(0, "صف الملف المقابل", unmatched.index + 2)
        unmatched.insert(1, "اتجاه المطابقة", direction)

    return results, unmatched


# -------------------------
# تقرير Excel منسق
# -------------------------

def make_report(sheets):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in sheets.items():
            if data is None:
                continue
            if not isinstance(data, pd.DataFrame):
                data = pd.DataFrame(data)
            data.to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False
            )

        workbook = writer.book

        header_fill = PatternFill(
            "solid", fgColor="6B244A"
        )
        header_font = Font(
            color="FFFFFF", bold=True, name="Arial"
        )
        normal_font = Font(name="Arial", size=10)

        for ws in workbook.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.rightToLeft = True

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )

            ws.row_dimensions[1].height = 32

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.font = normal_font
                    cell.alignment = Alignment(
                        horizontal="right",
                        vertical="center",
                        wrap_text=True
                    )

            for col_idx, column_cells in enumerate(
                ws.columns, start=1
            ):
                values = [
                    len(str(cell.value or ""))
                    for cell in list(column_cells)[:150]
                ]
                width = min(max(max(values, default=10) + 2, 12), 35)
                ws.column_dimensions[
                    get_column_letter(col_idx)
                ].width = width

    output.seek(0)
    return output


# -------------------------
# واجهة التطبيق
# -------------------------

st.header("1. رفع الملفين")

col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader(
        "الملف الأول",
        type=["xlsx", "xls"],
        key="upload1"
    )

with col2:
    file2 = st.file_uploader(
        "الملف الثاني",
        type=["xlsx", "xls"],
        key="upload2"
    )


if file1 and file2:
    try:
        df1 = read_uploaded_excel(file1, "الأول")
        df2 = read_uploaded_excel(file2, "الثاني")

        st.success(
            f"تم تحميل الملفين: {len(df1)} صفاً في الأول "
            f"و{len(df2)} صفاً في الثاني."
        )

        st.header("2. إعداد المطابقة")

        mode = st.radio(
            "نوع المطابقة",
            [
                "مطابقة دقيقة بالمفاتيح (مثل VLOOKUP)",
                "مطابقة ذكية متعددة الحقول"
            ],
            horizontal=True
        )

        direction = st.selectbox(
            "اتجاه المطابقة",
            [
                "الملف الأول مقابل الثاني",
                "الملف الثاني مقابل الأول",
                "المطابقة في الاتجاهين"
            ]
        )

        if mode == "مطابقة دقيقة بالمفاتيح (مثل VLOOKUP)":
            st.info(
                "اختر عموداً أو أكثر من كل ملف. "
                "يجب أن يكون ترتيب الأعمدة المختارة متقابلاً."
            )

            a, b = st.columns(2)

            with a:
                keys1 = st.multiselect(
                    "مفاتيح الملف الأول",
                    list(df1.columns),
                    key="exact_keys1"
                )

            with b:
                keys2 = st.multiselect(
                    "مفاتيح الملف الثاني",
                    list(df2.columns),
                    key="exact_keys2"
                )

            output1 = st.multiselect(
                "أعمدة الملف الأول في التقرير",
                list(df1.columns),
                default=list(df1.columns),
                key="exact_out1"
            )

            output2 = st.multiselect(
                "أعمدة الملف الثاني في التقرير",
                list(df2.columns),
                default=list(df2.columns),
                key="exact_out2"
            )

            if st.button(
                "ابدأ المطابقة الدقيقة",
                type="primary",
                key="exact_start"
            ):
                if not keys1 or not keys2:
                    st.error("اختر مفاتيح المطابقة من الملفين.")
                elif len(keys1) != len(keys2):
                    st.error("عدد أعمدة المفاتيح يجب أن يكون متساوياً.")
                else:
                    directions = []

                    if direction in [
                        "الملف الأول مقابل الثاني",
                        "المطابقة في الاتجاهين"
                    ]:
                        directions.append(
                            (df1, df2, keys1, keys2,
                             output2, "الأول مقابل الثاني")
                        )

                    if direction in [
                        "الملف الثاني مقابل الأول",
                        "المطابقة في الاتجاهين"
                    ]:
                        directions.append(
                            (df2, df1, keys2, keys1,
                             output1, "الثاني مقابل الأول")
                        )

                    all_results = []
                    all_unmatched = []

                    for source, target, sk, tk, out_cols, label in directions:
                        result, unmatched = exact_match(
                            source, target, sk, tk,
                            out_cols, label
                        )
                        all_results.append(result)
                        all_unmatched.append(unmatched)

                    results = pd.concat(
                        all_results, ignore_index=True
                    )
                    unmatched = pd.concat(
                        all_unmatched, ignore_index=True
                    )

                    duplicate1 = find_duplicates(df1, keys1)
                    duplicate2 = find_duplicates(df2, keys2)

                    summary = pd.DataFrame({
                        "المؤشر": [
                            "عدد صفوف الملف الأول",
                            "عدد صفوف الملف الثاني",
                            "عدد السجلات المتطابقة",
                            "عدد المفاتيح الفارغة",
                            "عدد المفاتيح المكررة في الأول",
                            "عدد المفاتيح المكررة في الثاني",
                            "عدد السجلات غير المقابلة"
                        ],
                        "العدد": [
                            len(df1),
                            len(df2),
                            (results["حالة المطابقة"] == "متطابق").sum(),
                            (results["حالة المطابقة"] == "مفتاح فارغ").sum(),
                            len(duplicate1),
                            len(duplicate2),
                            len(unmatched)
                        ]
                    })

                    st.session_state["report_sheets"] = {
                        "ملخص التحليل": summary,
                        "كل النتائج": results,
                        "غير المقابلة": unmatched,
                        "تكرار الملف الأول": duplicate1,
                        "تكرار الملف الثاني": duplicate2
                    }

        else:
            st.info(
                "اختر الحقول التي تريد استخدامها. "
                "يمكن اختيار حقل واحد أو عدة حقول."
            )

            fields = {}

            for field in WEIGHTS:
                a, b = st.columns(2)

                with a:
                    ca = st.selectbox(
                        f"{field} - الملف الأول",
                        ["بدون اختيار"] + list(df1.columns),
                        key=f"smart_a_{field}"
                    )

                with b:
                    cb = st.selectbox(
                        f"{field} - الملف الثاني",
                        ["بدون اختيار"] + list(df2.columns),
                        key=f"smart_b_{field}"
                    )

                if ca != "بدون اختيار" and cb != "بدون اختيار":
                    fields[field] = (ca, cb)

            threshold = st.slider(
                "الحد الأدنى لنسبة المطابقة",
                50, 100, 70
            )

            one_to_one = st.checkbox(
                "عدم استخدام السجل المقابل لأكثر من مرة",
                value=True
            )

            output1 = st.multiselect(
                "أعمدة الملف الأول في التقرير",
                list(df1.columns),
                default=list(df1.columns),
                key="smart_out1"
            )

            output2 = st.multiselect(
                "أعمدة الملف الثاني في التقرير",
                list(df2.columns),
                default=list(df2.columns),
                key="smart_out2"
            )

            if st.button(
                "ابدأ المطابقة الذكية",
                type="primary",
                key="smart_start"
            ):
                if not fields:
                    st.error("اختر حقلاً واحداً على الأقل.")
                else:
                    directions = []

                    if direction in [
                        "الملف الأول مقابل الثاني",
                        "المطابقة في الاتجاهين"
                    ]:
                        directions.append(
                            (df1, df2, fields, output2,
                             "الأول مقابل الثاني")
                        )

                    if direction in [
                        "الملف الثاني مقابل الأول",
                        "المطابقة في الاتجاهين"
                    ]:
                        reverse_fields = {
                            field: (pair[1], pair[0])
                            for field, pair in fields.items()
                        }
                        directions.append(
                            (df2, df1, reverse_fields, output1,
                             "الثاني مقابل الأول")
                        )

                    all_results = []
                    all_unmatched = []

                    for source, target, selected_fields, out_cols, label in directions:
                        result, unmatched = smart_match(
                            source,
                            target,
                            selected_fields,
                            out_cols,
                            threshold,
                            one_to_one,
                            label
                        )
                        all_results.append(result)
                        all_unmatched.append(unmatched)

                    results = pd.concat(
                        all_results, ignore_index=True
                    )
                    unmatched = pd.concat(
                        all_unmatched, ignore_index=True
                    )

                    summary = pd.DataFrame({
                        "المؤشر": [
                            "عدد صفوف الملف الأول",
                            "عدد صفوف الملف الثاني",
                            "تطابق مرتفع جداً",
                            "تطابق مرتفع - مراجعة",
                            "يحتاج مراجعة",
                            "غير موجود",
                            "غير المقابل في الاتجاه المحدد"
                        ],
                        "العدد": [
                            len(df1),
                            len(df2),
                            (results["حالة المطابقة"] == "تطابق مرتفع جداً").sum(),
                            (results["حالة المطابقة"] == "تطابق مرتفع - مراجعة").sum(),
                            (results["حالة المطابقة"] == "يحتاج مراجعة").sum(),
                            (results["حالة المطابقة"] == "غير موجود").sum(),
                            len(unmatched)
                        ]
                    })

                    st.session_state["report_sheets"] = {
                        "ملخص التحليل": summary,
                        "كل النتائج": results,
                        "غير المقابلة": unmatched
                    }

        # -------------------------
        # عرض النتائج وتنزيل التقرير
        # -------------------------

        if "report_sheets" in st.session_state:
            st.divider()
            st.header("3. نتائج المطابقة")

            sheets = st.session_state["report_sheets"]
            summary = sheets["ملخص التحليل"]
            results = sheets["كل النتائج"]

            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True
            )

            st.subheader("تفاصيل النتائج")
            st.dataframe(
                results,
                use_container_width=True,
                hide_index=True
            )

            if "غير المقابلة" in sheets:
                st.subheader(
                    "السجلات الموجودة في الملف المقابل وغير المرتبطة"
                )
                st.dataframe(
                    sheets["غير المقابلة"],
                    use_container_width=True,
                    hide_index=True
                )

            for name in ["تكرار الملف الأول", "تكرار الملف الثاني"]:
                if name in sheets:
                    st.subheader(name)
                    st.dataframe(
                        sheets[name],
                        use_container_width=True,
                        hide_index=True
                    )

            report = make_report(sheets)

            st.download_button(
                "📥 تنزيل تقرير المطابقة Excel",
                data=report,
                file_name="Smart_Excel_Matching_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.warning(
                "المطابقة الذكية تقديرية وليست إثباتاً نهائياً. "
                "راجع الحالات الملتبسة يدوياً قبل اعتماد النتائج."
            )

    except Exception as error:
        st.error(
            "حدث خطأ أثناء قراءة الملفات أو المطابقة. "
            f"تفاصيل الخطأ: {error}"
        )


st.divider()
st.caption(
    "Smart Excel Matcher | أداة لمقارنة ومطابقة ملفات Excel"
)
