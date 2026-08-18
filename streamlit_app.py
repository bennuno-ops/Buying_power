from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from calc import (
    CONSUMPTION_TO_PRICE,
    FEATURE_CHOICES,
    FEATURE_OPTIONS,
    NO_CHOICE,
    DataStore,
    create_result_workbook,
)

st.set_page_config(page_title="חישוב מדד רמת החיים האישי", page_icon="💰", layout="wide")

# Basic RTL styling for Hebrew content
st.markdown(
    """
    <style>
    .block-container { direction: rtl; text-align: right; }
    div[data-testid="stMetric"] { direction: rtl; text-align: right; }
    label, .stSelectbox label, .stNumberInput label { direction: rtl; text-align: right; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_data_store() -> DataStore:
    return DataStore()


try:
    data = get_data_store()
except Exception as exc:  # noqa: BLE001
    st.error(f"שגיאה בטעינת קובץ הנתונים: {exc}")
    st.stop()

st.title("חישוב מדד רמת החיים האישי")
st.caption("יוצרים להב ברנדל ועומר בן נון")

st.divider()
st.header("באיזה עשירון אנחנו?")

col1, col2 = st.columns(2)
with col1:
    income = st.number_input(
        "כמה נטו אתם מכניסים יחד בחודש? (כולל הכנסות פסיביות כמו הכנסה מדירה או מעסק צדדי)",
        min_value=0.0,
        step=100.0,
        value=0.0,
    )
with col2:
    persons = st.number_input("כמה נפשות אתם?", min_value=1, step=1, value=1)

if st.button("חשב את העשירון שלי", type="primary"):
    if income <= 0:
        st.error("יש להזין מספר חיובי")
    else:
        try:
            decile, group, income_per_standard = data.classify_decile(income, int(persons))
            st.success(
                f"אתם בעשירון {decile} וקבוצת העשירונים שמתאימה לכם היא קבוצת {group}\n\n"
                f"הכנסה חודשית נטו לנפש תקנית: {income_per_standard:,.1f} ש״ח"
            )
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))

st.divider()
st.header("חישוב מדד כוח הקניה של הקבוצה שלי")
st.subheader("בחירת מאפייני הקבוצה")

feature_cols = st.columns(3)
filters: dict[str, str] = {}
selected_features: list[str] = []
for idx, col in enumerate(feature_cols):
    with col:
        feature = st.selectbox(f"מאפיין {idx + 1}", FEATURE_CHOICES, index=len(FEATURE_CHOICES) - 1, key=f"feature_{idx}")
        selected_features.append(feature)
        if feature != NO_CHOICE:
            value = st.selectbox("בחר מהאפשרויות", FEATURE_OPTIONS[feature], key=f"value_{idx}")
            filters[feature] = value

if st.button("חשב את מדד כוח הקניה שלי", type="primary"):
    real_features = [f for f in selected_features if f != NO_CHOICE]
    if not real_features:
        st.error("לא נבחרו מאפייני קבוצה")
    elif len(real_features) != len(set(real_features)):
        st.error("לא ניתן לבחור את אותו מאפיין יותר מפעם אחת")
    else:
        try:
            result = data.build_persona_result(filters)
            st.session_state["last_result"] = result
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))
            st.session_state.pop("last_result", None)

result = st.session_state.get("last_result")
if result is not None:
    annual = result["annual"]
    summary = " | ".join(f"{k}: {v}" for k, v in result["filters"].items())
    st.subheader("מדד כוח הקניה לקבוצה:")
    st.write(summary)

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.markdown("**מדד כוח הקניה**")
        st.line_chart(annual.set_index("שנה")["מדד כוח קניה 2014=100"])
    with chart_col2:
        st.markdown("**מדד המחירים לצרכן המותאם לקבוצה שלכם**")
        st.line_chart(annual.set_index("שנה")["מדד מחירים אישי 2014=100"])

    display_df = annual[["שנה", "מדד כוח קניה 2014=100", "מדד מחירים אישי 2014=100"]].copy()
    display_df["שנה"] = display_df["שנה"].astype(int)
    display_df["מדד כוח קניה 2014=100"] = display_df["מדד כוח קניה 2014=100"].round(1)
    display_df["מדד מחירים אישי 2014=100"] = display_df["מדד מחירים אישי 2014=100"].round(1)
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    strongest_section, strongest_value = result["strongest_section"]
    st.write(
        f"סעיף ההוצאות שהכי השפיע עליך הוא: **{strongest_section}** "
        f"({strongest_value * 100:.1f} נקודות אחוז תרומה מצטברת בקירוב)"
    )
    st.caption(
        "בשנים 2024-2026 אין בטבלה מיקרו-דאטה חדש של משקי בית; הרכב הצריכה מחושב לפי השנה האחרונה "
        "שבה היו נתונים לקבוצה שלכם, ועליית השכר לפי העלייה הכללית במשק."
    )

    buffer = io.BytesIO()
    tmp_path = create_result_workbook(result, out_path=None)
    with open(tmp_path, "rb") as f:
        buffer.write(f.read())
    st.download_button(
        "הורד קובץ אקסל",
        data=buffer.getvalue(),
        file_name=tmp_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.caption("*המידע שאתם מכניסים לא נשמר, לא נאסף ולא חשוף לאף אחד, גם לא לנו.")
