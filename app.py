
import io
import re
import unicodedata
from collections import defaultdict

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

st.set_page_config(
    page_title="المطابقة الذكية لملفات Excel",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
html, body, [class*="css"] {
    direction: rtl;
    text-align: right;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 نظام المطابقة الذكية لملفات Excel")
st.caption("مقارنة وتحليل بيانات ملفي Excel باستخدام المطابقة الذكية")

st.info(
    "النتائج إرشادية وتحتاج مراجعة بشرية، "
    "خصوصاً عند وجود اختلافات في أرقام الهوية والحساب."
)


# تنظيف الأسماء العربية
def clean_name(value):
    if pd.isna(value):
        return ""

    value = unicodedata.normalize("NFKC", str(value))
    value = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", value)

    value = value.translate(str.maketrans({
        "أ": "ا", "إ": "ا", "آ": "ا",
        "ى": "ي", "ؤ": "و", "ئ": "ي"
    }))

    value = re.sub(r"\s+", " ", value.strip())

    return value.casefold()


# تنظيف الأرقام مع الحفاظ على الأصفار في البداية
def clean_number(value):
    if pd.isna(value):
        return ""

    value = str(value).strip()

    value = value.translate(str.maketrans(
        "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹",
        "01234567890123456789"
    ))

    value = re.sub(r"[\s\-_/.,]", "", value)

    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]

    return value


def get_value(row, column):
    value = row[column]
    return "" if pd.isna(value) else str(value).strip()


def similarity(a, b):
    if not a or not b:
        return None

    return max(
        fuzz.ratio(a, b),
        fuzz.token_sort_ratio(a, b)
    )


# محرك البحث عن السجلات المرشحة
def compare_files(df1, df2, c1, c2):

    names2 = [
        clean_name(get_value(df2.iloc[j], c2["name"]))
        for j in range(len(df2))
    ]

    name_choices = {
        j: name for j, name in enumerate(names2) if name
    }

    id_index = defaultdict(list)
    account_index = defaultdict(list)

    for j in range(len(df2)):

        ident = clean_number(
            get_value(df2.iloc[j], c2["id"])
        )

        account = clean_number(
            get_value(df2.iloc[j], c2["account"])
        )

        if ident:
            id_index[ident].append(j)

        if account:
            account_index[account].append(j)

    candidates_by_row = []
    all_candidates = []

    for i in range(len(df1)):

        r1 = df1.iloc[i]

        name1 = clean_name(get_value(r1, c1["name"]))
        type1 = clean_name(get_value(r1, c1["type"]))
        id1 = clean_number(get_value(r1, c1["id"]))
        account1 = clean_number(get_value(r1, c1["account"]))

        pool = set()

        if id1:
            pool.update(id_index.get(id1, []))

        if account1:
            pool.update(account_index.get(account1, []))

        if name1 and name_choices:
            found = process.extract(
                name1,
                name_choices,
                scorer=fuzz.WRatio,
                limit=10
            )

            for item in found:
                if item[1] >= 45:
                    pool.add(item[2])

        row_candidates = []

        for j in pool:

            r2 = df2.iloc[j]

            name2 = clean_name(get_value(r2, c2["name"]))
            type2 = clean_name(get_value(r2, c2["type"]))
            id2 = clean_number(get_value(r2, c2["id"]))
            account2 = clean_number(get_value(r2, c2["account"]))

            name_score = similarity(name1, name2)

            type_score = (
                100 if type1 and type2 and type1 == type2
                else 0 if type1 and type2
                else None
            )

            id_score = (
                100 if id1 and id2 and id1 == id2
                else 0 if id1 and id2
                else None
            )

            account_score = (
                100 if account1 and account2 and account1 == account2
                else 0 if account1 and account2
                else None
            )

            fields = [
                (name_score, 0.40),
                (type_score, 0.10),
                (id_score, 0.30),
                (account_score, 0.20)
            ]

            available = [
                (score, weight)
                for score, weight in fields
                if score is not None
            ]

            total = sum(weight for _, weight in available)

            score = (
                sum(value * weight for value, weight in available) / total
                if total else 0
            )

            candidate = {
                "source": i,
                "target": j,
                "score": round(score, 2),
                "name_score": round(name_score, 2)
                if name_score is not None else "",
                "type_match": "متطابق" if type_score == 100
                else "مختلف" if type_score == 0 else "غير متاح",
                "id_match": "متطابق" if id_score == 100
                else "مختلف" if id_score == 0 else "غير متاح",
                "account_match": "متطابق" if account_score == 100
                else "مختلف" if account_score == 0 else "غير متاح",
                "exact_id": bool(id1 and id2 and id1 == id2),
                "exact_account": bool(
                    account1 and account2 and account1 == account2
                )
            }

            row_candidates.append(candidate)
            all_candidates.append(candidate)

        row_candidates.sort(
            key=lambda x: (
                x["score"],
                x["exact_id"],
                x["exact_account"]
            ),
            reverse=True
        )

        candidates_by_row.append(row_candidates)

    return candidates_by_row, all_candidates


def make_results(
    df1, df2, c1, c2,
    candidates, all_candidates,
    one_to_one, threshold
):

    selected = {}

    if one_to_one:

        used = set()

        for candidate in sorted(
            all_candidates,
            key=lambda x: (
                x["score"],
                x["exact_id"],
                x["exact_account"]
            ),
            reverse=True
        ):

            if (
                candidate["source"] not in selected
                and candidate["target"] not in used
            ):

                selected[candidate["source"]] = candidate
                used.add(candidate["target"])

    else:

        for i, options in enumerate(candidates):
            if options:
                selected[i] = options[0]

    rows = []

    for i in range(len(df1)):

        r1 = df1.iloc[i]
        candidate = selected.get(i)

        row = {
            "رقم صف الملف الأول": i + 2,
            "اسم الملف الأول": get_value(r1, c1["name"]),
            "نوع الهوية - الأول": get_value(r1, c1["type"]),
            "رقم الهوية - الأول": get_value(r1, c1["id"]),
            "رقم الحساب - الأول": get_value(r1, c1["account"])
        }

        if candidate:

            r2 = df2.iloc[candidate["target"]]
            score = candidate["score"]

            status = (
                "تطابق مرتفع جداً" if score >= 95
                else "تطابق مرتفع" if score >= 85
                else "يحتاج مراجعة" if score >= 70
                else "تطابق ضعيف"
            )

            row.update({
                "رقم صف الملف الثاني": candidate["target"] + 2,
                "اسم الملف الثاني": get_value(r2, c2["name"]),
                "نوع الهوية - الثاني": get_value(r2, c2["type"]),
                "رقم الهوية - الثاني": get_value(r2, c2["id"]),
                "رقم الحساب - الثاني": get_value(r2, c2["account"]),
                "نسبة التطابق %": score,
                "تشابه الاسم %": candidate["name_score"],
                "مطابقة نوع الهوية": candidate["type_match"],
                "مطابقة رقم الهوية": candidate["id_match"],
                "مطابقة رقم الحساب": candidate["account_match"],
                "حالة المطابقة": status,
                "ملاحظة": "يرجى مراجعة الاختلافات"
                if score < threshold else ""
            })

        else:

            row.update({
                "رقم صف الملف الثاني": "",
                "اسم الملف الثاني": "",
                "نوع الهوية - الثاني": "",
                "رقم الهوية - الثاني": "",
                "رقم الحساب - الثاني": "",
                "نسبة التطابق %": 0,
                "تشابه الاسم %": "",
                "مطابقة نوع الهوية": "غير متاح",
                "مطابقة رقم الهوية": "غير متاح",
                "مطابقة رقم الحساب": "غير متاح",
                "حالة المطابقة": "غير موجود في الملف الثاني",
                "ملاحظة": "لم يتم العثور على سجل مرشح"
            })

        rows.append(row)

    result = pd.DataFrame(rows)

    matched = result[
        result["رقم صف الملف الثاني"] != ""
    ]

    review = result[
        result["حالة المطابقة"].isin(
            ["يحتاج مراجعة", "تطابق ضعيف"]
        )
    ]

    unmatched = result[
        result["حالة المطابقة"] == "غير موجود في الملف الثاني"
    ]

    summary = pd.DataFrame({
        "المؤشر": [
            "سجلات الملف الأول",
            "سجلات الملف الثاني",
            "السجلات التي عثر لها على مقابل",
            "السجلات غير الموجودة",
            "الحالات التي تحتاج مراجعة",
            "المطابقات بنسبة 95% فأكثر",
            "متوسط نسبة التطابق"
        ],
        "القيمة": [
            len(df1),
            len(df2),
            len(matched),
            len(unmatched),
            len(review),
            int((matched["نسبة التطابق %"] >= 95).sum()),
            round(matched["نسبة التطابق %"].mean(), 2)
            if len(matched) else 0
        ]
    })

    return result, summary, review, unmatched


# واجهة البرنامج

st.sidebar.header("إعدادات المطابقة")

threshold = st.sidebar.slider(
    "حد التنبيه للمراجعة",
    50, 100, 85
)

one_to_one = st.sidebar.checkbox(
    "عدم تكرار سجل الملف الثاني",
    value=True
)

st.subheader("1. رفع ملفي Excel")

col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader(
        "الملف الأول",
        type=["xlsx", "xls"]
    )

with col2:
    file2 = st.file_uploader(
        "الملف الثاني",
        type=["xlsx", "xls"]
    )


if file1 and file2:

    try:

        df1 = pd.read_excel(file1, dtype=str).fillna("")
        df2 = pd.read_excel(file2, dtype=str).fillna("")

        st.success(
            f"الملف الأول: {len(df1):,} سجل | "
            f"الملف الثاني: {len(df2):,} سجل"
        )

        st.subheader("2. تحديد أعمدة المقارنة")

        cols1 = list(df1.columns)
        cols2 = list(df2.columns)

        mapping1 = {}
        mapping2 = {}

        fields = {
            "name": ("الاسم", ["اسم", "name"]),
            "type": ("نوع الهوية", ["نوع الهوية", "نوع الوثيقة", "id type"]),
            "id": ("رقم الهوية", ["رقم الهوية", "رقم الوثيقة", "id number"]),
            "account": ("رقم الحساب", ["رقم الحساب", "account"])
        }

        for field, (label, keywords) in fields.items():

            a, b = st.columns(2)

            def guess(columns):
                for column in columns:
                    if any(k in clean_name(column) for k in keywords):
                        return column
                return columns[0]

            with a:
                default = guess(cols1)
                mapping1[field] = st.selectbox(
                    label + " - الملف الأول",
                    cols1,
                    index=cols1.index(default),
                    key="first_" + field
                )

            with b:
                default = guess(cols2)
                mapping2[field] = st.selectbox(
                    label + " - الملف الثاني",
                    cols2,
                    index=cols2.index(default),
                    key="second_" + field
                )

        if st.button(
            "🚀 بدء المطابقة والتحليل",
            type="primary",
            use_container_width=True
        ):

            with st.spinner("جارٍ تحليل البيانات..."):

                candidates, all_candidates = compare_files(
                    df1, df2, mapping1, mapping2
                )

                result, summary, review, unmatched = make_results(
                    df1, df2,
                    mapping1, mapping2,
                    candidates, all_candidates,
                    one_to_one, threshold
                )

                st.session_state["data"] = (
                    result, summary, review, unmatched
                )

    except Exception as error:
        st.error(f"حدث خطأ أثناء قراءة الملفات: {error}")


if "data" in st.session_state:

    result, summary, review, unmatched = st.session_state["data"]

    st.divider()
    st.subheader("3. نتائج التحليل")

    values = dict(
        zip(summary["المؤشر"], summary["القيمة"])
    )

    a, b, c, d = st.columns(4)

    a.metric("سجلات الملف الأول", values["سجلات الملف الأول"])
    b.metric("تم العثور على مقابل", values["السجلات التي عثر لها على مقابل"])
    c.metric("غير موجود", values["السجلات غير الموجودة"])
    d.metric("متوسط التطابق", f'{values["متوسط نسبة التطابق"]}%')

    tab1, tab2, tab3, tab4 = st.tabs([
        "كل النتائج",
        "الملخص",
        "يحتاج مراجعة",
        "غير موجود"
    ])

    with tab1:
        st.dataframe(result, use_container_width=True)

    with tab2:
        st.dataframe(summary, use_container_width=True)

    with tab3:
        st.dataframe(review, use_container_width=True)

    with tab4:
        st.dataframe(unmatched, use_container_width=True)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        summary.to_excel(
            writer, index=False, sheet_name="ملخص التحليل"
        )

        result.to_excel(
            writer, index=False, sheet_name="كل النتائج"
        )

        review.to_excel(
            writer, index=False, sheet_name="يحتاج مراجعة"
        )

        unmatched.to_excel(
            writer, index=False, sheet_name="غير موجود"
        )

        for ws in writer.book.worksheets:

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for column in ws.columns:
                width = max(
                    len(str(cell.value or ""))
                    for cell in column[:200]
                )
                ws.column_dimensions[
                    column[0].column_letter
                ].width = min(width + 3, 40)

    st.download_button(
        "📥 تنزيل تقرير Excel الكامل",
        data=output.getvalue(),
        file_name="تقرير_المطابقة_الذكية.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

else:

    st.warning("ارفع الملفين وحدد الأعمدة لبدء التحليل.")

st.divider()
st.caption("Smart Excel Matcher | الإصدار الأول 1.0")
