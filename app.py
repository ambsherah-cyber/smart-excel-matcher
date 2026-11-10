
import streamlit as st
import pandas as pd
import re
import io
import unicodedata
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
.stApp { direction: rtl; }
</style>
""", unsafe_allow_html=True)

st.title("📊 برنامج المطابقة الذكية لملفات Excel")
st.caption("مطابقة دقيقة بالمفاتيح المحددة أو مطابقة ذكية متعددة الحقول")

WEIGHTS = {
    "الاسم": 40,
    "نوع الهوية": 10,
    "رقم الهوية": 30,
    "رقم الحساب": 20
}


def normalize(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    value = unicodedata.normalize("NFKC", value)
    value = re.sub(r"[\u064B-\u065F\u0670]", "", value)
    value = value.replace("ـ", "")
    value = value.translate(str.maketrans(
        "أإآىؤئ",
        "ااايوي"
    ))
    value = re.sub(r"\s+", " ", value)
    return value.lower().strip()


def normalize_number(value):
    value = normalize(value)
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    persian_digits = "۰۱۲۳۴۵۶۷۸۹"
    english_digits = "0123456789"
    value = value.translate(str.maketrans(arabic_digits, english_digits))
    value = value.translate(str.maketrans(persian_digits, english_digits))
    value = re.sub(r"\s+", "", value)
    return value


def read_excel(uploaded_file):
    excel = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox(
        f"اختر ورقة العمل في {uploaded_file.name}",
        excel.sheet_names,
        key=f"sheet_{uploaded_file.name}"
    )
    return pd.read_excel(uploaded_file, sheet_name=sheet, dtype=str).fillna("")


def make_key(row, columns):
    values = []
    for col in columns:
        values.append(normalize_number(row[col]))
    return "||".join(values)


def similarity(a, b):
    if not a or not b:
        return None
    return fuzz.ratio(a, b)


def score_rows(row1, row2, fields):
    total = 0
    used_weight = 0
    details = {}

    for field, columns in fields.items():
        col1, col2 = columns
        a = normalize(row1[col1])
        b = normalize(row2[col2])

        if not a or not b:
            continue

        if field in ["رقم الهوية", "رقم الحساب"]:
            a = normalize_number(row1[col1])
            b = normalize_number(row2[col2])

        value = fuzz.ratio(a, b)
        weight = WEIGHTS[field]
        total += value * weight
        used_weight += weight
        details[field] = round(value, 1)

    final_score = total / used_weight if used_weight else 0
    return round(final_score, 1), details


def category(score):
    if score >= 95:
        return "تطابق مرتفع جداً"
    if score >= 85:
        return "تطابق مرتفع"
    if score >= 70:
        return "يحتاج مراجعة"
    return "تطابق ضعيف"


def to_excel(results, summary, duplicates1=None, duplicates2=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="ملخص التحليل", index=False)
        results.to_excel(writer, sheet_name="كل النتائج", index=False)

        review = results[
            results["حالة المطابقة"].astype(str).str.contains(
                "مراجعة|مكرر|ضعيف", na=False
            )
        ]
        review.to_excel(writer, sheet_name="يحتاج مراجعة", index=False)

        missing = results[
            results["حالة المطابقة"].astype(str).str.contains(
                "غير موجود", na=False
            )
        ]
        missing.to_excel(writer, sheet_name="غير موجود", index=False)

        if duplicates1 is not None:
            duplicates1.to_excel(writer, sheet_name="تكرار الملف الأول", index=False)
        if duplicates2 is not None:
            duplicates2.to_excel(writer, sheet_name="تكرار الملف الثاني", index=False)

    output.seek(0)
    return output


st.markdown("### 1. رفع ملفات Excel")

c1, c2 = st.columns(2)
with c1:
    file1 = st.file_uploader("الملف الأول", type=["xlsx", "xls"], key="file1")
with c2:
    file2 = st.file_uploader("الملف الثاني", type=["xlsx", "xls"], key="file2")

if file1 and file2:
    try:
        df1 = read_excel(file1)
        df2 = read_excel(file2)

        st.success(
            f"تم تحميل الملفين بنجاح: "
            f"{len(df1)} صفاً في الملف الأول، و{len(df2)} صفاً في الملف الثاني."
        )

        st.markdown("### 2. اختر طريقة المطابقة")

        mode = st.radio(
            "طريقة المطابقة",
            [
                "المطابقة بالمفتاح المحدد (مثل VLOOKUP)",
                "المطابقة الذكية متعددة الحقول"
            ],
            horizontal=False
        )

        st.markdown("### 3. إعداد أعمدة المطابقة")

        if mode == "المطابقة بالمفتاح المحدد (مثل VLOOKUP)":
            st.info(
                "اختر عموداً واحداً أو عدة أعمدة من كل ملف. "
                "يجب أن يكون ترتيب الأعمدة المختارة متقابلاً."
            )

            k1, k2 = st.columns(2)
            with k1:
                key_cols1 = st.multiselect(
                    "مفاتيح المطابقة من الملف الأول",
                    list(df1.columns),
                    key="key_cols1"
                )
            with k2:
                key_cols2 = st.multiselect(
                    "مفاتيح المطابقة من الملف الثاني",
                    list(df2.columns),
                    key="key_cols2"
                )

            output_cols = st.multiselect(
                "الأعمدة المطلوب جلبها من الملف الثاني",
                list(df2.columns),
                default=list(df2.columns)
            )

            if st.button("ابدأ المطابقة", type="primary", key="run_key"):
                if not key_cols1 or not key_cols2:
                    st.error("يرجى اختيار أعمدة المطابقة من الملفين.")
                elif len(key_cols1) != len(key_cols2):
                    st.error("يجب اختيار العدد نفسه من الأعمدة في الملفين.")
                else:
                    keys1 = df1.apply(lambda r: make_key(r, key_cols1), axis=1)
                    keys2 = df2.apply(lambda r: make_key(r, key_cols2), axis=1)

                    duplicate_keys1 = keys1[keys1.duplicated(keep=False)]
                    duplicate_keys2 = keys2[keys2.duplicated(keep=False)]

                    duplicate_report1 = df1.loc[duplicate_keys1.index].copy()
                    duplicate_report2 = df2.loc[duplicate_keys2.index].copy()

                    counts2 = keys2.value_counts()
                    lookup = {}
                    for idx, key in enumerate(keys2):
                        lookup.setdefault(key, []).append(idx)

                    records = []
                    for idx, row in df1.iterrows():
                        key = keys1.loc[idx]
                        matches = lookup.get(key, [])

                        record = {
                            "صف الملف الأول": idx + 2,
                            "مفتاح المطابقة": key
                        }

                        for col in df1.columns:
                            record[f"الملف الأول - {col}"] = row[col]

                        if not key.replace("|", "").strip():
                            status = "مفتاح فارغ"
                            match_idx = None
                        elif len(matches) > 1:
                            status = "مكرر في الملف الثاني"
                            match_idx = None
                        elif len(matches) == 1:
                            status = "متطابق"
                            match_idx = matches[0]
                        else:
                            status = "غير موجود في الملف الثاني"
                            match_idx = None

                        record["حالة المطابقة"] = status
                        record["صف الملف الثاني"] = (
                            matches[0] + 2 if len(matches) == 1 else ""
                        )

                        for col in output_cols:
                            record[f"الملف الثاني - {col}"] = (
                                df2.iloc[match_idx][col]
                                if match_idx is not None else ""
                            )

                        records.append(record)

                    results = pd.DataFrame(records)

                    summary = pd.DataFrame({
                        "المؤشر": [
                            "عدد صفوف الملف الأول",
                            "عدد صفوف الملف الثاني",
                            "السجلات المتطابقة",
                            "غير الموجودة",
                            "المكررة في الملف الثاني",
                            "المفاتيح الفارغة",
                            "المفاتيح المكررة في الملف الأول"
                        ],
                        "العدد": [
                            len(df1),
                            len(df2),
                            (results["حالة المطابقة"] == "متطابق").sum(),
                            (results["حالة المطابقة"] == "غير موجود في الملف الثاني").sum(),
                            (results["حالة المطابقة"] == "مكرر في الملف الثاني").sum(),
                            (results["حالة المطابقة"] == "مفتاح فارغ").sum(),
                            len(duplicate_report1)
                        ]
                    })

                    st.session_state["results"] = results
                    st.session_state["summary"] = summary
                    st.session_state["dup1"] = duplicate_report1
                    st.session_state["dup2"] = duplicate_report2

        else:
            st.info(
                "اختر الحقول المتاحة في كل ملف. "
                "يمكن ترك أي حقل غير متوفر دون اختيار."
            )

            fields = {}
            for field in WEIGHTS:
                st.markdown(f"**{field}**")
                a, b = st.columns(2)
                with a:
                    col_a = st.selectbox(
                        f"{field} - الملف الأول",
                        ["— بدون اختيار —"] + list(df1.columns),
                        key=f"a_{field}"
                    )
                with b:
                    col_b = st.selectbox(
                        f"{field} - الملف الثاني",
                        ["— بدون اختيار —"] + list(df2.columns),
                        key=f"b_{field}"
                    )

                if col_a != "— بدون اختيار —" and col_b != "— بدون اختيار —":
                    fields[field] = (col_a, col_b)

            output_cols = st.multiselect(
                "الأعمدة المطلوب إظهارها من الملف الثاني",
                list(df2.columns),
                default=list(df2.columns),
                key="smart_output"
            )

            threshold = st.slider(
                "الحد الأدنى لقبول المرشح للمطابقة",
                min_value=50,
                max_value=100,
                value=70
            )

            one_to_one = st.checkbox(
                "منع استخدام سجل الملف الثاني لأكثر من سجل",
                value=True
            )

            if st.button("ابدأ المطابقة الذكية", type="primary", key="run_smart"):
                if not fields:
                    st.error("يرجى اختيار حقل واحد على الأقل من الملفين.")
                else:
                    name_cols = fields.get("الاسم")
                    if name_cols:
                        name_values = [
                            normalize(v) for v in df2[name_cols[1]].tolist()
                        ]
                    else:
                        name_values = []

                    used = set()
                    records = []

                    for idx, row in df1.iterrows():
                        candidates = set()

                        # البحث عن مرشحين عبر القيم الدقيقة للهوية والحساب
                        for field in ["رقم الهوية", "رقم الحساب"]:
                            if field in fields:
                                c1, c2 = fields[field]
                                value = normalize_number(row[c1])
                                if value:
                                    matches = df2.index[
                                        df2[c2].apply(normalize_number) == value
                                    ].tolist()
                                    candidates.update(matches)

                        # البحث التقريبي بالاسم
                        if name_cols:
                            query = normalize(row[name_cols[0]])
                            if query and name_values:
                                found = process.extract(
                                    query,
                                    name_values,
                                    scorer=fuzz.ratio,
                                    limit=10
                                )
                                for _, value_score, pos in found:
                                    if value_score >= 40:
                                        candidates.add(df2.index[pos])

                        # عند عدم وجود مرشح واضح، افحص كل السجلات إذا كان الملف صغيراً
                        if not candidates and len(df2) <= 3000:
                            candidates = set(df2.index.tolist())

                        best_idx = None
                        best_score = -1
                        best_details = {}

                        for candidate_idx in candidates:
                            if one_to_one and candidate_idx in used:
                                continue

                            candidate = df2.loc[candidate_idx]
                            score, details = score_rows(row, candidate, fields)

                            if score > best_score:
                                best_score = score
                                best_idx = candidate_idx
                                best_details = details

                        record = {"صف الملف الأول": idx + 2}
                        for col in df1.columns:
                            record[f"الملف الأول - {col}"] = row[col]

                        if best_idx is None or best_score < threshold:
                            record["نسبة التطابق"] = (
                                best_score if best_idx is not None else 0
                            )
                            record["تصنيف التطابق"] = (
                                category(best_score) if best_idx is not None else "غير موجود"
                            )
                            record["حالة المطابقة"] = (
                                "يحتاج مراجعة" if best_idx is not None else "غير موجود"
                            )
                            record["صف الملف الثاني"] = ""
                            for col in output_cols:
                                record[f"الملف الثاني - {col}"] = ""
                        else:
                            used.add(best_idx)
                            record["نسبة التطابق"] = best_score
                            record["تصنيف التطابق"] = category(best_score)
                            record["حالة المطابقة"] = (
                                "متطابق" if best_score >= 95 else
                                "يحتاج مراجعة"
                            )
                            record["صف الملف الثاني"] = best_idx + 2

                            for col in output_cols:
                                record[f"الملف الثاني - {col}"] = df2.loc[best_idx, col]

                        for field in WEIGHTS:
                            record[f"نسبة {field}"] = best_details.get(field, "")

                        records.append(record)

                    results = pd.DataFrame(records)

                    summary = pd.DataFrame({
                        "المؤشر": [
                            "عدد صفوف الملف الأول",
                            "عدد صفوف الملف الثاني",
                            "تطابق مرتفع جداً",
                            "تطابق مرتفع",
                            "يحتاج مراجعة",
                            "غير موجود"
                        ],
                        "العدد": [
                            len(df1),
                            len(df2),
                            (results["نسبة التطابق"] >= 95).sum(),
                            ((results["نسبة التطابق"] >= 85) &
                             (results["نسبة التطابق"] < 95)).sum(),
                            ((results["نسبة التطابق"] >= threshold) &
                             (results["نسبة التطابق"] < 85)).sum(),
                            (results["حالة المطابقة"] == "غير موجود").sum()
                        ]
                    })

                    st.session_state["results"] = results
                    st.session_state["summary"] = summary
                    st.session_state["dup1"] = None
                    st.session_state["dup2"] = None

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملفات أو إعداد المطابقة: {e}")


if "results" in st.session_state:
    st.markdown("---")
    st.markdown("## 4. نتائج المطابقة")

    results = st.session_state["results"]
    summary = st.session_state["summary"]

    a, b, c = st.columns(3)
    a.metric("إجمالي سجلات الملف الأول", len(results))
    b.metric(
        "السجلات المتطابقة",
        (results["حالة المطابقة"] == "متطابق").sum()
    )
    c.metric(
        "السجلات غير الموجودة",
        results["حالة المطابقة"].astype(str).str.contains("غير موجود").sum()
    )

    st.dataframe(results, use_container_width=True, hide_index=True)

    st.markdown("### ملخص التحليل")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    report = to_excel(
        results,
        summary,
        st.session_state.get("dup1"),
        st.session_state.get("dup2")
    )

    st.download_button(
        "📥 تنزيل تقرير المطابقة Excel",
        data=report,
        file_name="تقرير_المطابقة.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.warning(
        "نسب التشابه مؤشرات مساعدة وليست إثباتاً نهائياً لهوية المستفيد. "
        "راجع الحالات الملتبسة يدوياً قبل اعتماد أي إجراء."
    )

st.markdown("---")
st.caption("Smart Excel Matcher | أداة مساعدة لمقارنة ملفات Excel")
