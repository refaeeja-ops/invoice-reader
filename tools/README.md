# تدريب مصنّف الأرقام الهندية

1. نزّل خطوطًا عربية مجانية من Google Fonts إلى مجلد واحد، وعدّل مسار `glob` في `gen.py`.
2. `pip install numpy pillow scipy fonttools`
3. `python3 gen.py train 48000 && python3 gen.py test 4000 && python3 train.py`
4. انسخ `model.json` الناتج إلى `vendor/digits-model.json`.

`seg.py` يطابق خوارزمية التقطيع في `digits.js` حرفيًا؛ أي تعديل في أحدهما يجب أن ينعكس على الآخر.
النتيجة الحالية: 98.8% دقة لكل رمز على خطين لم يتدرّب عليهما (Noto Naskh و Tajawal).
