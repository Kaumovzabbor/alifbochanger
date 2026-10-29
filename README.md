# Alifbo Konverter

Word (.docx) fayllaridagi o'zbek matnini **eski lotin alifbosidan yangi o'zbek lotin alifbosiga** o'tkazadigan Windows ilovasi.

| Eski | Yangi |
|---|---|
| oʻ / Oʻ | ö / Ö |
| gʻ / Gʻ | ğ / Ğ |
| sh / Sh / SH | ş / Ş |
| ch / Ch / CH | ç / Ç |

"ng", tutuq belgisi (ta’lim, a’zo, ma’no) va boshqa harflar o'zgarmaydi. Faylning dizayni, shriftlari, ranglari, rasmlari, jadvallari, matn maydonlari, kolontitullari va formatlashi saqlanadi. Dastur internetga ulanmaydi, hujjat matni hech qayerga yuborilmaydi va logga yozilmaydi.

---

## 1. Foydalanish (tayyor .exe bilan)

1. `AlifboKonverter.exe` ni ishga tushiring.
2. Fayllarni oynaga **tashlang** (yoki «Fayl tanlash» / «Papka tanlash»). Papka ichidagi hamma .docx fayllar (ichki papkalar bilan) qo'shiladi.
3. Kerak bo'lsa «Natija papkasi» ni tanlang (bo'sh bo'lsa, natija asl fayl yonida saqlanadi).
4. **«Yangi alifboga o'tkazish»** tugmasini bosing.
5. Natija `asl_nom_yangi_alifbo.docx` nomi bilan saqlanadi. **Asl fayl o'zgarmaydi.** Nom band bo'lsa `_2`, `_3` qo'shiladi.
6. Tugagach hisobot chiqadi, «Natija papkasini ochish» tugmasi yoqiladi.

**Ko'rib chiqish:** fayllarni o'zgartirmasdan, qaysi so'zlar o'zgarishini ko'rsatadi (eski → yangi, soni bilan). O'zgarmasligi kerak so'zlarni belgilab «O'zgartirma» ni bosing, ular istisnolarga qo'shiladi.

**Sozlamalar:** tutuq belgisi turi, til filtri, kirill → lotin rejimi, natija nomi qo'shimchasi, asl faylni almashtirish (`.bak` nusxa bilan), stop-list, istisnolar, tema.

### Til filtri
«Faqat o'zbekcha deb belgilangan matnni o'zgartirish» yoqilganda (standart), Word'da **ingliz, rus va boshqa tilga** belgilangan matn bo'laklari (run) o'tkazib yuboriladi. Tili belgilanmagan yoki `uz` deb belgilangan matn o'zgartiriladi. Agar hujjatingiz butunlay o'zbekcha bo'lsa-yu, matn adashib «English» deb belgilangan bo'lsa, shu tugmani o'chirib qo'ying.

### O'zgartirilmaydigan narsalar
Raqamlar, URL, e-pochta, `C:\yo'l\fayl`, fayl nomlari (`hisobot.pdf`), domenlar (`maktab.uz`), maydon kodlari (field), xatcho'plar, stil nomlari, XML tuzilma. Stop-listdagi so'zlar (masalan `Chelsea`, `Shell`) ham o'zgarmaydi.

### Istisnolar
Sozlamalarda har qatorga `eski => yangi` yozing. Masalan, o yoki g dan keyin **haqiqiy tutuq** kelgan so'z: `mo'tabar => mo’tabar`. «O'zgartirma» deb belgilash uchun `so'z => so'z` (ko'rib chiqish oynasi buni o'zi qiladi). Katta-kichik harf avtomatik moslanadi. Fayl joyi: `%APPDATA%\AlifboKonverter\exceptions.json`.

### Interfeys matnlarini almashtirish
Hamma tugma va yozuvlar `app\strings_uz.json` da. O'z variantingizni `%APPDATA%\AlifboKonverter\strings.json` ga yozsangiz (faqat o'zgargan kalitlar yetarli), u standart matnlarni almashtiradi.

### Buyruq satri
```
AlifboKonverter.exe fayl.docx -o papka
AlifboKonverter.exe C:\Hujjatlar --preview
AlifboKonverter.exe fayl.docx --replace --apostrophe modifier --all-languages
AlifboKonverter.exe --help
```
Argument berilmasa oyna ochiladi. Eslatma: `--windowed` yig'ilgan .exe natijani chaqirilgan konsolga yozadi (cmd/PowerShell), ba'zan chiqish kechikib ko'rinishi mumkin.

---

## 2. .exe ni yig'ish (Python o'rnatilmagan kompyuterda ham ishlaydi)

Yig'ish uchun **Windows 10/11 + Python 3.11 yoki yangiroq** kerak (python.org dan o'rnating, o'rnatishda «Add python.exe to PATH» belgisini qo'ying).

Loyiha papkasida `cmd` oching:

```
py -3.11 -m venv venv
venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m pytest
build.bat
```

Tayyor fayl: `dist\AlifboKonverter.exe` (bitta fayl, Python kerak emas). `python -m pytest` hamma testlar o'tganini ko'rsatishi kerak; avval shuni tekshirib oling.

Dasturni yig'masdan sinash: `python main.py` (oyna) yoki `python main.py fayl.docx` (buyruq satri).

### GitHub Actions orqali yig'ish (Windows kompyuter kerak emas)

1. GitHub'da yangi repozitoriy yarating va shu papkadagi hamma fayllarni (`.github` papkasi bilan) yuklang.
2. **Actions** bo'limiga kiring: «Windows .exe yig'ish» avtomatik ishga tushadi (yoki **Run workflow** tugmasini bosing).
3. Tugagach (taxminan 3-5 daqiqa) ish sahifasining pastidagi **Artifacts** dan yuklab oling:
   - `AlifboKonverter-exe`: bitta `AlifboKonverter.exe`
   - `AlifboKonverter-papka`: papka varianti (antivirus ogohlantirsa shuni ishlating)
4. Workflow o'zi testlarni o'tkazadi va yig'ilgan .exe ni haqiqatan ishga tushirib sinaydi; biror narsa buzilsa, artifact yaratilmaydi.
5. Doimiy havola kerak bo'lsa, teg qo'ying: `git tag v1.0.0 && git push origin v1.0.0`. Fayllar **Releases** bo'limiga chiqadi.

### Antivirus / Windows Defender ogohlantirishi
PyInstaller bilan yig'ilgan bitta-fayl .exe uchun antivirus ba'zan **soxta ogohlantirish** beradi. Yechimlar:
- `build_onedir.bat` bilan yig'ing: bitta fayl o'rniga papka (`dist\AlifboKonverter\`) hosil bo'ladi, ogohlantirish kamroq bo'ladi.
- Dasturni o'zingiz yig'ganingizni tasdiqlash uchun Defender'da papkani istisnolarga qo'shing.
- Keng tarqatish uchun .exe ni kod imzolash sertifikati bilan imzolang (`signtool sign ...`).

---

## 3. Loyiha tuzilmasi

```
alifbo_konverter/
├── app/
│   ├── converter.py        sof alifbo mantig'i (convert_text, convert_spans)
│   ├── cyrillic.py         ixtiyoriy: kirill -> yangi lotin
│   ├── docx_processor.py   zip + lxml: .docx ni o'qish, run'larga qaytarish, yozish
│   ├── batch.py            ko'p fayllarni ishlash (GUI va CLI uchun umumiy)
│   ├── settings.py         sozlamalar/istisnolar (%APPDATA%)
│   ├── strings.py, strings_uz.json   interfeys matnlari
│   ├── exceptions.json     standart istisnolar (bo'sh)
│   ├── gui.py              oyna
│   └── cli.py              buyruq satri
├── tests/                  pytest testlari
├── assets/icon.ico
├── main.py
├── requirements.txt
├── build.bat, build_onedir.bat
└── README.md
```

## 4. Qanday ishlaydi (qisqacha)

Word bitta so'zni bir nechta bo'lakka (run) bo'lib yozishi mumkin (`bo` + `’` + `limi`). Shuning uchun har bir paragrafdagi matn bitta satrga birlashtiriladi, alifbo qoidalari shu satrga qo'llanadi, so'ng har bir almashtirish tegishli bo'lakka qaytariladi. Har bir bo'lak o'z formatini (qalin, rang, shrift) saqlaydi. Tab, qator uzilishi, maydon belgilari va rasmlar so'z chegarasi hisoblanadi. O'zgarmagan fayl qismlari bayt-ma-bayt nusxalanadi; natija yozilgach o'zi tekshiriladi (zip butunligi, XML to'g'riligi, o'zgarmasligi kerak qismlar).

Qayta ishlanadigan joylar: asosiy matn, jadvallar, matn maydonlari (ikkala `mc:AlternateContent` varianti ham), kolontitullar, izohlar (footnotes/endnotes/comments), diagramma (chart) va SmartArt matni, rasmning muqobil matni, kuzatilgan o'zgarishlar (`w:ins`/`w:del`), hujjat xossalari (sarlavha, mavzu, kalit so'zlar).

## 5. Cheklovlar

- Parol bilan himoyalangan va eski `.doc` fayllar qo'llab-quvvatlanmaydi (tushunarli xabar ko'rsatiladi va keyingi faylga o'tiladi). Ularni Word'da `.docx` qilib saqlang.
- Diagramma (chart) ichiga biriktirilgan Excel jadvali (`word/embeddings`) o'zgartirilmaydi; Word'da diagrammani «Ma'lumotlarni tahrirlash» orqali ochsangiz, eski yozuvlar ko'rinishi mumkin.
- Til filtri faqat run'ning o'zida belgilangan tilni (`w:lang`) hisobga oladi, stil yoki hujjat standart tilini emas.
- Ingliz tilidagi matnda til belgilanmagan bo'lsa, `sh`/`ch` va `'` o'zbekcha deb hisoblanadi; bunday so'zlarni stop-listga qo'shing.
- Natija papkasi tanlansa, papkalar tuzilmasi ko'chirilmaydi: hamma natija bitta papkaga yoziladi.
