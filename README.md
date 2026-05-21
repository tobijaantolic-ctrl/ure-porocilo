# Poročilo opravljenih ur — Setup

Web app ki pretvori PDF poročilo učitelja v Excel z razčlenjenimi urami.

## Enkratni setup (~10 minut)

### 1. Gemini API ključ (brezplačno)
1. Pojdi na https://aistudio.google.com/app/apikey
2. Klikni **Create API key**
3. Kopiraj ključ — shrani ga za korak 4

### 2. GitHub repo
1. Pojdi na https://github.com/new
2. Ime: `ure-porocilo` (ali karkoli)
3. Pusti **Public**
4. Klikni **Create repository**
5. Naloži te 3 datoteke: `app.py`, `requirements.txt`, `README.md`

### 3. Streamlit Community Cloud (brezplačno)
1. Pojdi na https://share.streamlit.io
2. Prijavi se z GitHub računom
3. Klikni **New app**
4. Izberi tvoj repo in branch `main`, file `app.py`
5. Klikni **Advanced settings → Secrets**

### 4. Dodaj API ključ v Secrets
V polju Secrets vnesi točno to:
```
GEMINI_API_KEY = "tukaj-tvoj-kljuc"
```

### 5. Deploy
Klikni **Deploy** — čez ~2 minuti dobiš URL tipa:
`https://ure-porocilo.streamlit.app`

Ta URL pošlješ vsem učiteljem. Nobene prijave, nobene instalacije.

## Uporaba
1. Odpri URL
2. Naloži PDF poročilo
3. Klikni "Obdelaj"
4. Prenesi Excel

## Stroški
- Streamlit Cloud: **brezplačno**
- Gemini API: **brezplačno** do 1500 klicev/dan
- Za 50 učiteljev 2x/leto = 100 klicev/leto → daleč pod limitom
