# מדד רמת החיים האישי — Streamlit App

גרסת Streamlit של מחשבון מדד כוח הקניה האישי, בנויה על אותה לוגיקת חישוב
מהתוכנה המקומית (Tkinter), עטופה בממשק ווב.

## מבנה התיקייה

- `streamlit_app.py` — ממשק המשתמש (Streamlit).
- `calc.py` — לוגיקת החישוב (זהה לתוכנה המקומית, ללא תלות ב-Tkinter).
- `outputs/living_standard_panel_2012_2026/...xlsx` — קובץ נתוני הבסיס.
- `requirements.txt` — התלויות הדרושות.

## הרצה מקומית

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

האפליקציה תיפתח בדפדפן בכתובת `http://localhost:8501`.

## פריסה ל-Streamlit Community Cloud

1. מעלים את כל תוכן התיקייה הזו ל-repository חדש ב-GitHub (כולל תיקיית
   `outputs/`).
2. נכנסים ל-[share.streamlit.io](https://share.streamlit.io) ומתחברים עם
   חשבון GitHub.
3. לוחצים על **New app**, בוחרים את ה-repository, את הענף (`main`), ואת
   הקובץ הראשי: `streamlit_app.py`.
4. לוחצים **Deploy**. אחרי דקה-שתיים האתר יהיה זמין בכתובת ציבורית.

## פרטיות

הקלט שהמשתמש מזין אינו נשמר בשום מקום קבוע — כל חישוב מתבצע בזיכרון של
הסשן בלבד, וקובץ האקסל שנוצר מוצג להורדה ישירה מהדפדפן.
