# -*- coding: utf-8 -*-
"""
Script to compile all 26 Tamyiz Huruf / Kalimat columns into:
1. Kompilasi_Metode_Tamyiz_26_Kolom.docx
2. Kompilasi_Metode_Tamyiz_26_Kolom.html
3. Kompilasi_Metode_Tamyiz_26_Kolom.md
"""

import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

COLUMNS_DATA = [
    {
        "nomor": 1,
        "judul_arab": "بِيَجُرّ",
        "judul_latin": "Bi-yajur (Huruf Jer)",
        "kategori": "Huruf Jer",
        "kaidah": "Huruf yang membarisbawahkan (mengkashrahkan) kata benda/isim setelahnya.",
        "items": [
            {"arab": "بِ...", "latin": "bi", "arti": "dengan", "freq": "-"},
            {"arab": "كَ...", "latin": "ka", "arti": "seperti", "freq": "-"},
            {"arab": "لِ... / لَ...", "latin": "li / la", "arti": "untuk", "freq": "-"},
            {"arab": "إِلَى / إِلَيْهِ", "latin": "ilaa / ilaihi", "arti": "kepada", "freq": "434"},
            {"arab": "عَلَى / عَلَيْهِ", "latin": "'alaa / 'alaihi", "arti": "di atas", "freq": "1.445"},
            {"arab": "مِنْ", "latin": "min", "arti": "dari", "freq": "2.299"},
            {"arab": "فِيْ / فِيْهِ", "latin": "fii / fiihi", "arti": "di dalam", "freq": "1.699"},
            {"arab": "عَنْ...", "latin": "'an", "arti": "dari", "freq": "155"},
            {"arab": "حَتَّى", "latin": "hattaa", "arti": "sehingga", "freq": "-"},
            {"arab": "بِ... تَ... وَ...", "latin": "bi, ta, wa (Qasam)", "arti": "demi", "freq": "-"}
        ]
    },
    {
        "nomor": 2,
        "judul_arab": "كَانَ رَفْعُ نَصَبَ",
        "judul_latin": "Kaana Rafa' Nashab",
        "kategori": "Kaana wa Akhawatuha",
        "kaidah": "Kata kerja bantu yang merofa'kan isim (subjek) dan menashabkan khabar (predikat).",
        "items": [
            {"arab": "كَانَ", "latin": "kaana", "arti": "adalah", "freq": "1.418"},
            {"arab": "لَيْسَ", "latin": "laisa", "arti": "tiada, bukan", "freq": "79"}
        ]
    },
    {
        "nomor": 3,
        "judul_arab": "إِنَّ نَصَبَ رَفْعُ",
        "judul_latin": "Inna Nashab Rafa'",
        "kategori": "Inna wa Akhawatuha",
        "kaidah": "Huruf penegas yang menashabkan isim dan merofa'kan khabar.",
        "items": [
            {"arab": "إِنَّ", "latin": "inna", "arti": "sesungguhnya", "freq": "972"},
            {"arab": "أَنَّ", "latin": "anna", "arti": "sesungguhnya", "freq": "187"},
            {"arab": "كَأَنَّ", "latin": "ka-anna", "arti": "seakan-akan", "freq": "37"},
            {"arab": "لَكِنَّ", "latin": "lakinna", "arti": "akan tetapi", "freq": "65"},
            {"arab": "لَعَلَّ", "latin": "la'alla", "arti": "semoga, supaya", "freq": "129"},
            {"arab": "لَيْتَ", "latin": "laita", "arti": "andaikan", "freq": "14"}
        ]
    },
    {
        "nomor": 4,
        "judul_arab": "لَا نَصَبَ لِلنَّكِرَةِ",
        "judul_latin": "Laa Nashab lin-Nakirah",
        "kategori": "Laa Nafiyah lil Jinsi",
        "kaidah": "Huruf peniadaan mutlak terhadap seluruh jenis, menashabkan kata umum (nakirah).",
        "items": [
            {"arab": "لَا", "latin": "laa", "arti": "tidak ada", "freq": "1.731"}
        ]
    },
    {
        "nomor": 5,
        "judul_arab": "إِلَّا نَصَبَ لِلْمُسْتَثْنَاءِ",
        "judul_latin": "Illaa Nashab lil-Mustatsna'",
        "kategori": "Huruf Istitsna'",
        "kaidah": "Huruf pengecualian yang menashabkan kata yang dikecualikan (mustatsna).",
        "items": [
            {"arab": "إِلَّا", "latin": "illaa", "arti": "kecuali", "freq": "470"}
        ]
    },
    {
        "nomor": 6,
        "judul_arab": "يَا نَصَبَ لِلْمُضَافِ",
        "judul_latin": "Yaa Nashab lil-Mudhaf",
        "kategori": "Yaa Munada Mudhaf",
        "kaidah": "Huruf panggilan yang menashabkan kata majemuk bersandar (mudhaf).",
        "items": [
            {"arab": "يَا", "latin": "yaa", "arti": "wahai", "freq": "112"}
        ]
    },
    {
        "nomor": 7,
        "judul_arab": "يَا نِدَاءُ",
        "judul_latin": "Yaa Nida' (Panggilan)",
        "kategori": "Huruf Nida'",
        "kaidah": "Kata seru/panggilan untuk menarik perhatian orang yang diseru.",
        "items": [
            {"arab": "يَا", "latin": "yaa", "arti": "wahai", "freq": "112"},
            {"arab": "يَا أَيُّهَا", "latin": "yaa ayyuhaa", "arti": "wahai (laki-laki / umum)", "freq": "156"},
            {"arab": "يَا أَيَّتُهَا", "latin": "yaa ayyatuhaa", "arti": "wahai (perempuan)", "freq": "1"}
        ]
    },
    {
        "nomor": 8,
        "judul_arab": "أَنْ يَنْصِبَ",
        "judul_latin": "An Yanshiba (Penashab Fi'il)",
        "kategori": "Nawasib al-Mudhari'",
        "kaidah": "Huruf yang membarisfathahkan (menashabkan) kata kerja mudhari'.",
        "items": [
            {"arab": "أَنْ", "latin": "an", "arti": "hendaknya", "freq": "577"},
            {"arab": "لَنْ", "latin": "lan", "arti": "tidak akan", "freq": "104"},
            {"arab": "إِذَنْ", "latin": "idzan", "arti": "kalau demikian", "freq": "3"},
            {"arab": "كَيْ", "latin": "kay", "arti": "supaya", "freq": "10"},
            {"arab": "لِ...", "latin": "li", "arti": "supaya", "freq": "-"},
            {"arab": "حَتَّى", "latin": "hattaa", "arti": "hingga", "freq": "71"}
        ]
    },
    {
        "nomor": 9,
        "judul_arab": "لَا تَجْزُمْ",
        "judul_latin": "Laa Tajzum (Larangan)",
        "kategori": "Laa an-Nahiyah",
        "kaidah": "Huruf larangan yang mensukunkan (menjazamkan) fi'il mudhari'.",
        "items": [
            {"arab": "لَا", "latin": "laa", "arti": "janganlah", "freq": "1.731"}
        ]
    },
    {
        "nomor": 10,
        "judul_arab": "لَمْ يَجْزُمْ",
        "judul_latin": "Lam Yajzum (Penjazam Fi'il)",
        "kategori": "Jawazim al-Mudhari'",
        "kaidah": "Huruf peniadaan dan perintah yang menjazamkan (mensukunkan) satu fi'il mudhari'.",
        "items": [
            {"arab": "لَمْ", "latin": "lam", "arti": "tidak / belum", "freq": "347"},
            {"arab": "لَمَّا", "latin": "lammaa", "arti": "tidak / belum", "freq": "165"},
            {"arab": "لِ...", "latin": "li", "arti": "hendaklah", "freq": "-"},
            {"arab": "وَلْ...", "latin": "wal", "arti": "hendaklah", "freq": "9"},
            {"arab": "فَلْ...", "latin": "fal", "arti": "hendaklah", "freq": "7"}
        ]
    },
    {
        "nomor": 11,
        "judul_arab": "الشَّرْطُ وَالْجَوَابُ",
        "judul_latin": "Asy-Syarthu wal Jawabu",
        "kategori": "Jawazim Syarthiyah",
        "kaidah": "Kata-kata syarat yang menjazamkan dua kata kerja sekaligus (fi'il syarat dan fi'il jawab).",
        "items": [
            {"arab": "إِنْ", "latin": "in", "arti": "jika", "freq": "628"},
            {"arab": "مَنْ", "latin": "man", "arti": "siapa", "freq": "274"},
            {"arab": "مَا", "latin": "maa", "arti": "apa", "freq": "718"},
            {"arab": "أَيْنَمَا", "latin": "ainamaa", "arti": "kemanapun", "freq": "4"},
            {"arab": "حَيْثُمَا", "latin": "haitsumaa", "arti": "dimanapun", "freq": "1"},
            {"arab": "مَهْمَا", "latin": "mahmaa", "arti": "bagaimanapun juga", "freq": "1"}
        ]
    },
    {
        "nomor": 12,
        "judul_arab": "اَلْإِسْتِثْنَاءُ",
        "judul_latin": "Al-Istitsna'",
        "kategori": "Istitsna' & Taghrid",
        "kaidah": "Kata-kata pengecualian, peringatan keras, atau pencegahan.",
        "items": [
            {"arab": "إِلَّا", "latin": "illaa", "arti": "kecuali", "freq": "470"},
            {"arab": "أَلَّا", "latin": "allaa", "arti": "hendaklah janganlah", "freq": "45"},
            {"arab": "كَلَّا", "latin": "kallaa", "arti": "sekali-kali janganlah", "freq": "33"}
        ]
    },
    {
        "nomor": 13,
        "judul_arab": "اَلْعَطَفُ",
        "judul_latin": "Al-'Athafu (Kata Sambung)",
        "kategori": "Huruf 'Athaf",
        "kaidah": "Kata penghubung antar kata atau kalimat yang menyamakan kedudukan i'rab.",
        "items": [
            {"arab": "وَ...", "latin": "wa", "arti": "dan", "freq": "-"},
            {"arab": "أَوْ", "latin": "au", "arti": "atau", "freq": "265"},
            {"arab": "فَ...", "latin": "fa", "arti": "maka", "freq": "-"},
            {"arab": "أَمْ", "latin": "am", "arti": "atau", "freq": "122"},
            {"arab": "ثُمَّ", "latin": "tsumma", "arti": "kemudian", "freq": "338"},
            {"arab": "أَمَّا", "latin": "ammaa", "arti": "adapun", "freq": "59"},
            {"arab": "بَلْ...", "latin": "bal", "arti": "bahkan", "freq": "127"},
            {"arab": "حَتَّى...", "latin": "hattaa", "arti": "hingga", "freq": "71"},
            {"arab": "لَكِنْ", "latin": "laakin", "arti": "akan tetapi", "freq": "65"}
        ]
    },
    {
        "nomor": 14,
        "judul_arab": "اَلشَّرْطُ",
        "judul_latin": "Asy-Syarthu (Syarat Non-Jazam)",
        "kategori": "Adawat Syarthi Ghair Jazimah",
        "kaidah": "Kata-kata pengandaian, waktu terjadinya syarat, tanpa menjazamkan kata kerja.",
        "items": [
            {"arab": "إِذْ", "latin": "idz", "arti": "jika / ketika", "freq": "239"},
            {"arab": "إِذَا", "latin": "idzaa", "arti": "jika / apabila", "freq": "432"},
            {"arab": "إِذَنْ", "latin": "idzan", "arti": "jika / kalau begitu", "freq": "-"},
            {"arab": "إِمَّا", "latin": "immaa", "arti": "jika tidak", "freq": "29"},
            {"arab": "أَمَّا", "latin": "ammaa", "arti": "adapun", "freq": "59"},
            {"arab": "لَمَّا", "latin": "lammaa", "arti": "ketika", "freq": "165"},
            {"arab": "لَوْ", "latin": "lau", "arti": "jikalau", "freq": "200"},
            {"arab": "لَوْلَا", "latin": "laulaa", "arti": "mengapa tidak / jikalau tidak", "freq": "75"},
            {"arab": "لَوْمَا", "latin": "laumaa", "arti": "mengapa tidak", "freq": "1"}
        ]
    },
    {
        "nomor": 15,
        "judul_arab": "اَلْإِسْتِفْهَامُ",
        "judul_latin": "Al-Istifham (Kata Tanya)",
        "kategori": "Adawat al-Istifham",
        "kaidah": "Kata-kata tanya di dalam Al-Qur'an untuk menanyakan orang, benda, waktu, atau alasan.",
        "items": [
            {"arab": "أَ", "latin": "a", "arti": "apakah", "freq": "-"},
            {"arab": "أَلَا", "latin": "alaa", "arti": "ingatlah", "freq": "45"},
            {"arab": "أَيُّ", "latin": "ayyu", "arti": "siapakah / yang mana", "freq": "59"},
            {"arab": "أَيْنَ", "latin": "aina", "arti": "dimana", "freq": "7"},
            {"arab": "مَا", "latin": "maa", "arti": "apakah", "freq": "718"},
            {"arab": "مَنْ", "latin": "man", "arti": "siapakah", "freq": "274"},
            {"arab": "مَتَى", "latin": "mataa", "arti": "kapan", "freq": "9"},
            {"arab": "كَمْ", "latin": "kam", "arti": "berapa", "freq": "21"},
            {"arab": "كَيْفَ", "latin": "kaifa", "arti": "bagaimana", "freq": "83"},
            {"arab": "هَلْ", "latin": "hal", "arti": "apakah", "freq": "92"},
            {"arab": "مَاذَا", "latin": "maadzaa", "arti": "apakah", "freq": "26"},
            {"arab": "لِمَ", "latin": "lima", "arti": "kenapa", "freq": "19"},
            {"arab": "لِمَاذَا", "latin": "limaadzaa", "arti": "kenapa", "freq": "-"}
        ]
    },
    {
        "nomor": 16,
        "judul_arab": "اَلتَّوْكِيْدُ",
        "judul_latin": "At-Taukid (Penegasan)",
        "kategori": "Huruf Taukid",
        "kaidah": "Huruf penegas dan penguat makna kalimat (sungguh / benar-benar).",
        "items": [
            {"arab": "لَ...", "latin": "la", "arti": "sungguh", "freq": "-"},
            {"arab": "قَدْ", "latin": "qad", "arti": "sungguh", "freq": "218"},
            {"arab": "لَقَدْ", "latin": "laqad", "arti": "sungguh", "freq": "188"}
        ]
    },
    {
        "nomor": 17,
        "judul_arab": "اَلْإِسْتِقْبَالُ",
        "judul_latin": "Al-Istiqbal (Masa Depan)",
        "kategori": "Huruf Tanfis / Istiqbal",
        "kaidah": "Huruf yang masuk pada fi'il mudhari' untuk menentukan waktu yang akan datang (akan).",
        "items": [
            {"arab": "سَ...", "latin": "sa", "arti": "akan", "freq": "-"},
            {"arab": "سَوْفَ", "latin": "saufa", "arti": "akan", "freq": "42"}
        ]
    },
    {
        "nomor": 18,
        "judul_arab": "اَلنَّافِي",
        "judul_latin": "An-Nafi (Peniadaan)",
        "kategori": "Huruf Nafi",
        "kaidah": "Huruf peniadaan yang berarti 'bukan / tidak'.",
        "items": [
            {"arab": "مَا", "latin": "maa", "arti": "bukan", "freq": "718"},
            {"arab": "لَا", "latin": "laa", "arti": "bukan", "freq": "1.731"}
        ]
    },
    {
        "nomor": 19,
        "judul_arab": "[tanpa nama] / أَفْعَالُ الْمَدْحِ وَالذَّمِّ",
        "judul_latin": "Tanpa Nama (Pujian & Celaan)",
        "kategori": "Af'alul Madh wadz-Dzam",
        "kaidah": "Kata kerja khusus untuk memuji setinggi-tingginya atau mencela seburuk-buruknya.",
        "items": [
            {"arab": "نِعْمَ", "latin": "ni'ma", "arti": "sebaik-baik", "freq": "16"},
            {"arab": "بِئْسَ", "latin": "bi'sa", "arti": "sejelek-jelek", "freq": "40"}
        ]
    },
    {
        "nomor": 20,
        "judul_arab": "ظَرْفٌ",
        "judul_latin": "Zhorof (Keterangan Tempat & Waktu)",
        "kategori": "Zharaf Makan wa Zaman",
        "kaidah": "Kata keterangan penunjuk posisi ruang (tempat) atau urutan waktu.",
        "items": [
            {"arab": "قَبْلَ", "latin": "qabla", "arti": "sebelum", "freq": "188"},
            {"arab": "بَعْدَ", "latin": "ba'da", "arti": "sesudah", "freq": "199"},
            {"arab": "غَيْرَ", "latin": "ghaira", "arti": "bukan / selain", "freq": "147"},
            {"arab": "دُوْنَ", "latin": "duuna", "arti": "selain", "freq": "144"},
            {"arab": "أَمَامَ", "latin": "amaama", "arti": "di depan", "freq": "2"},
            {"arab": "وَرَاءَ", "latin": "waraa'a", "arti": "di belakang", "freq": "24"},
            {"arab": "خَلْفَ", "latin": "khalfa", "arti": "di belakang", "freq": "19"},
            {"arab": "فَوْقَ", "latin": "fauqa", "arti": "di atas", "freq": "41"},
            {"arab": "تَحْتَ", "latin": "tahta", "arti": "di bawah", "freq": "51"},
            {"arab": "جَانِبَ", "latin": "jaaniba", "arti": "di samping", "freq": "9"},
            {"arab": "حَوْلَ", "latin": "haula", "arti": "di sekitar", "freq": "17"},
            {"arab": "كُلَّ", "latin": "kulla", "arti": "setiap", "freq": "407"},
            {"arab": "مَعَ", "latin": "ma'a", "arti": "beserta", "freq": "59"},
            {"arab": "عِنْدَ", "latin": "'inda", "arti": "di sisi", "freq": "196"},
            {"arab": "بَيْنَ", "latin": "baina", "arti": "di antara", "freq": "266"}
        ]
    },
    {
        "nomor": 21,
        "judul_arab": "مَوْصُوْل",
        "judul_latin": "Maushul (Kata Sambung Orang / Benda)",
        "kategori": "Isim Maushul",
        "kaidah": "Kata penghubung (relative pronoun) 'yang / orang-orang yang' untuk laki-laki dan perempuan.",
        "items": [
            {"arab": "اَلَّذِيْ", "latin": "alladzii", "arti": "orang yang (lelaki 1)", "freq": "304"},
            {"arab": "اَللَّذَانِ", "latin": "alladzaani", "arti": "orang yang (lelaki 2)", "freq": "-"},
            {"arab": "اَلَّذِيْنَ", "latin": "alladziina", "arti": "orang-orang yang (lelaki banyak)", "freq": "1.080"},
            {"arab": "اَلَّتِيْ", "latin": "allatii", "arti": "orang yang (wanita 1)", "freq": "68"},
            {"arab": "اَلتَّانِ", "latin": "allataani", "arti": "orang yang (wanita 2)", "freq": "-"},
            {"arab": "اَللَّاتِيْ / اَللَّائِيْ", "latin": "allaatii / allaa'ii", "arti": "orang-orang yang (wanita banyak)", "freq": "-"},
            {"arab": "مَنْ", "latin": "man", "arti": "siapa yang", "freq": "274"},
            {"arab": "مَا", "latin": "maa", "arti": "apa yang", "freq": "718"}
        ]
    },
    {
        "nomor": 22,
        "judul_arab": "إِشَارَةٌ (بَعِيْد)",
        "judul_latin": "Isyarah Jauh (Kata Tunjuk 'Itu')",
        "kategori": "Isim Isyarah lil Ba'id",
        "kaidah": "Kata tunjuk untuk menunjuk benda/orang yang posisinya jauh ('itu / mereka itu').",
        "items": [
            {"arab": "ذٰلِكَ", "latin": "dzaalika", "arti": "itu (lelaki 1)", "freq": "427"},
            {"arab": "ذٰلِكُمَا", "latin": "dzaalikumaa", "arti": "itu (lelaki 2)", "freq": "2"},
            {"arab": "ذٰلِكُمْ", "latin": "dzaalikum", "arti": "itu (lelaki jamak)", "freq": "27"},
            {"arab": "تِلْكَ", "latin": "tilka", "arti": "itu (wanita 1)", "freq": "41"},
            {"arab": "تِلْكُمَا", "latin": "tilkumaa", "arti": "itu (wanita 2)", "freq": "-"},
            {"arab": "تِلْكُمْ", "latin": "tilkum", "arti": "itu (wanita jamak)", "freq": "-"},
            {"arab": "أُولٰئِكَ", "latin": "ulaaa'ika", "arti": "mereka itu", "freq": "204"}
        ]
    },
    {
        "nomor": 23,
        "judul_arab": "إِشَارَةٌ (قَرِيْب)",
        "judul_latin": "Isyarah Dekat (Kata Tunjuk 'Ini')",
        "kategori": "Isim Isyarah lil Qarib",
        "kaidah": "Kata tunjuk untuk menunjuk benda/orang yang posisinya dekat ('ini / mereka ini').",
        "items": [
            {"arab": "هٰذَا", "latin": "haadzaa", "arti": "ini satu (lelaki)", "freq": "225"},
            {"arab": "هٰذَانِ", "latin": "haadzaani", "arti": "ini dua (lelaki)", "freq": "2"},
            {"arab": "هٰؤُلَاءِ", "latin": "haa'ulaaa'i", "arti": "mereka ini (lelaki)", "freq": "46"},
            {"arab": "هٰذِهِ", "latin": "haadzihi", "arti": "ini satu (wanita)", "freq": "47"},
            {"arab": "هَاتَانِ", "latin": "haataani", "arti": "ini dua (wanita)", "freq": "2"},
            {"arab": "هٰؤُلَاءِ", "latin": "haa'ulaaa'i", "arti": "mereka ini (wanita)", "freq": "46"}
        ]
    },
    {
        "nomor": 24,
        "judul_arab": "ضَمِيْرٌ (مُنْفَصِلٌ مَرْفُوْعٌ)",
        "judul_latin": "Dhamir Munfashil (Kata Ganti Mandiri)",
        "kategori": "Dhamir Munfashil",
        "kaidah": "Kata ganti orang yang berdiri sendiri (tidak bersambung), berfungsi sebagai subjek.",
        "items": [
            {"arab": "هُوَ", "latin": "huwa", "arti": "dia (lelaki 1)", "freq": "-"},
            {"arab": "هُمَا", "latin": "humaa", "arti": "dia berdua (lelaki 2)", "freq": "-"},
            {"arab": "هُمْ", "latin": "hum", "arti": "mereka (lelaki banyak)", "freq": "-"},
            {"arab": "هِيَ", "latin": "hiya", "arti": "dia (wanita 1)", "freq": "-"},
            {"arab": "هُمَا", "latin": "humaa", "arti": "dia berdua (wanita 2)", "freq": "-"},
            {"arab": "هُنَّ", "latin": "hunna", "arti": "mereka (wanita banyak)", "freq": "-"},
            {"arab": "أَنْتَ", "latin": "anta", "arti": "kamu (lelaki 1)", "freq": "-"},
            {"arab": "أَنْتُمَا", "latin": "antumaa", "arti": "kamu berdua (lelaki 2)", "freq": "-"},
            {"arab": "أَنْتُمْ", "latin": "antum", "arti": "kalian (lelaki banyak)", "freq": "-"},
            {"arab": "أَنْتِ", "latin": "anti", "arti": "kamu (wanita 1)", "freq": "-"},
            {"arab": "أَنْتُمَا", "latin": "antumaa", "arti": "kamu berdua (wanita 2)", "freq": "-"},
            {"arab": "أَنْتُنَّ", "latin": "antunna", "arti": "kalian (wanita banyak)", "freq": "-"},
            {"arab": "أَنَا", "latin": "anaa", "arti": "saya", "freq": "-"},
            {"arab": "نَحْنُ", "latin": "nahnu", "arti": "kami / kita", "freq": "-"}
        ]
    },
    {
        "nomor": 25,
        "judul_arab": "ضَمِيْرٌ (مُتَّصِلٌ)",
        "judul_latin": "Dhamir Muttashil (Kata Ganti Bersambung)",
        "kategori": "Dhamir Muttashil",
        "kaidah": "Akhiran kata ganti yang menempel di belakang kata benda (milik) atau kata kerja (objek).",
        "items": [
            {"arab": "...ـهُ / ...ـهِ", "latin": "-hu / -hi", "arti": "dia / -nya (lelaki 1)", "freq": "-"},
            {"arab": "...ـهُمَا / ...ـهِمَا", "latin": "-humaa / -himaa", "arti": "mereka berdua (lelaki 2)", "freq": "-"},
            {"arab": "...ـهُمْ / ...ـهِمْ", "latin": "-hum / -him", "arti": "mereka (lelaki banyak)", "freq": "-"},
            {"arab": "...ـهَا", "latin": "-haa", "arti": "dia / -nya (wanita 1)", "freq": "-"},
            {"arab": "...ـهُمَا / ...ـهِمَا", "latin": "-humaa / -himaa", "arti": "mereka berdua (wanita 2)", "freq": "-"},
            {"arab": "...ـهُنَّ / ...ـهِنَّ", "latin": "-hunna / -hinna", "arti": "mereka (wanita banyak)", "freq": "-"},
            {"arab": "...ـكَ", "latin": "-ka", "arti": "kamu / -mu (lelaki 1)", "freq": "-"},
            {"arab": "...ـكُمَا", "latin": "-kumaa", "arti": "kalian berdua (lelaki 2)", "freq": "-"},
            {"arab": "...ـكُمْ", "latin": "-kum", "arti": "kalian (lelaki banyak)", "freq": "-"},
            {"arab": "...ـكِ", "latin": "-ki", "arti": "kamu / -mu (wanita 1)", "freq": "-"},
            {"arab": "...ـكُمَا", "latin": "-kumaa", "arti": "kalian berdua (wanita 2)", "freq": "-"},
            {"arab": "...ـكُنَّ", "latin": "-kunna", "arti": "kalian (wanita banyak)", "freq": "-"},
            {"arab": "...ـيْ / ...ـِيْ / ...ـنِيْ", "latin": "-ya / -ii / -nii", "arti": "saya / -ku", "freq": "-"},
            {"arab": "...ـنَا", "latin": "-naa", "arti": "kami / kita / -kita", "freq": "-"}
        ]
    },
    {
        "nomor": 26,
        "judul_arab": "ضَمِيْرٌ (مُنْفَصِلٌ مَنْصُوْبٌ)",
        "judul_latin": "Dhamir Munfashil Manshub (Kata Ganti Pengkhususan)",
        "kategori": "Dhamir Nashab Munfashil",
        "kaidah": "Kata ganti manshub pengkhususan pembatasan (Hasyr/Takhshish: 'Hanya kepada...').",
        "items": [
            {"arab": "إِيَّاهُ", "latin": "iyyaahu", "arti": "hanya kepadanya (lelaki 1)", "freq": "-"},
            {"arab": "إِيَّاهُمَا", "latin": "iyyaahumaa", "arti": "hanya kepadanya berdua (lelaki 2)", "freq": "-"},
            {"arab": "إِيَّاهُمْ", "latin": "iyyaahum", "arti": "hanya kepada mereka (lelaki banyak)", "freq": "-"},
            {"arab": "إِيَّاهَا", "latin": "iyyaahaa", "arti": "hanya kepadanya (wanita 1)", "freq": "-"},
            {"arab": "إِيَّاهُمَا", "latin": "iyyaahumaa", "arti": "hanya kepadanya berdua (wanita 2)", "freq": "-"},
            {"arab": "إِيَّاهُنَّ", "latin": "iyyaahunna", "arti": "hanya kepada mereka (wanita banyak)", "freq": "-"},
            {"arab": "إِيَّاكَ", "latin": "iyyaaka", "arti": "hanya kepadamu (lelaki 1)", "freq": "-"},
            {"arab": "إِيَّاكُمَا", "latin": "iyyaakumaa", "arti": "hanya kepadamu berdua (lelaki 2)", "freq": "-"},
            {"arab": "إِيَّاكُمْ", "latin": "iyyaakum", "arti": "hanya kepada kalian (lelaki banyak)", "freq": "-"},
            {"arab": "إِيَّاكِ", "latin": "iyyaaki", "arti": "hanya kepadamu (wanita 1)", "freq": "-"},
            {"arab": "إِيَّاكُمَا", "latin": "iyyaakumaa", "arti": "hanya kepadamu berdua (wanita 2)", "freq": "-"},
            {"arab": "إِيَّاكُنَّ", "latin": "iyyaakunna", "arti": "hanya kepada kalian (wanita banyak)", "freq": "-"},
            {"arab": "إِيَّايَ", "latin": "iyyaaya", "arti": "hanya kepadaku", "freq": "-"},
            {"arab": "إِيَّانَا", "latin": "iyyaanaa", "arti": "hanya kepada kami", "freq": "-"}
        ]
    }
]

def build_markdown(filepath):
    lines = []
    lines.append("# BUKU PANDUAN LENGKAP 26 KOLOM METODE TAMYIZ")
    lines.append("## Terjemah Huruf, Kata Penghubung & Kata Ganti dalam Al-Qur'an")
    lines.append("")
    lines.append("> **Catatan Pengantar:**")
    lines.append("> Materi ini dikompilasi dari 26 slide foto WhatsApp modul pembelajaran **Tamyiz Online (Metode Tamyiz)**.")
    lines.append("> Kata-kata dalam 26 kolom ini merupakan huruf dan kata bantu yang **paling sering berulang di seluruh ayat Al-Qur'an**.")
    lines.append("> Angka frekuensi menunjukkan perkiraan **jumlah kemunculan kata tersebut di dalam Mushaf Al-Qur'an**.")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## DAFTAR ISI 26 KOLOM")
    lines.append("")
    lines.append("| No | Nama Kolom | Arab | Kategori / Kaidah |")
    lines.append("|:---:|:---|:---:|:---|")
    for col in COLUMNS_DATA:
        lines.append(f"| {col['nomor']} | [{col['judul_latin']}](#kolom-{col['nomor']}) | {col['judul_arab']} | {col['kategori']} |")
    lines.append("")
    lines.append("---")
    lines.append("")

    for col in COLUMNS_DATA:
        lines.append(f"### <a id='kolom-{col['nomor']}'></a>Kolom {col['nomor']}: {col['judul_arab']} ({col['judul_latin']})")
        lines.append(f"**Kategori:** {col['kategori']}  ")
        lines.append(f"**Kaidah:** *{col['kaidah']}*  ")
        lines.append("")
        lines.append("| No | Lafadz Arab | Transliterasi | Terjemah Indonesia | Frekuensi Al-Qur'an |")
        lines.append("|:---:|:---:|:---|:---|:---:|")
        for idx, it in enumerate(col['items'], 1):
            freq_str = f"**{it['freq']}x**" if it['freq'] != "-" else "-"
            lines.append(f"| {idx} | <span style='font-size:1.3em;'>{it['arab']}</span> | {it['latin']} | {it['arti']} | {freq_str} |")
        lines.append("")
        lines.append("---")
        lines.append("")

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Markdown saved: {filepath}")

def build_html(filepath):
    html_content = """<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Buku Panduan 26 Kolom Huruf Metode Tamyiz - Terjemah Al-Qur'an</title>
    <link href="https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #0f766e;
            --primary-dark: #115e59;
            --primary-light: #ccfbf1;
            --accent: #f59e0b;
            --badge-bg: #fee2e2;
            --badge-color: #dc2626;
            --bg-body: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #1e293b;
            --text-muted: #64748b;
            --border-color: #e2e8f0;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: var(--bg-body);
            color: var(--text-main);
            line-height: 1.6;
            padding-bottom: 60px;
        }

        .header {
            background: linear-gradient(135deg, #0d9488 0%, #115e59 100%);
            color: white;
            padding: 40px 20px;
            text-align: center;
            box-shadow: 0 4px 20px rgba(13, 148, 136, 0.2);
            position: relative;
        }

        .header h1 {
            font-size: 2.2rem;
            font-weight: 800;
            letter-spacing: -0.5px;
            margin-bottom: 8px;
        }

        .header p {
            font-size: 1.1rem;
            opacity: 0.92;
            max-width: 750px;
            margin: 0 auto;
        }

        .header .badge-total {
            display: inline-block;
            background: rgba(255, 255, 255, 0.2);
            backdrop-filter: blur(8px);
            padding: 6px 16px;
            border-radius: 9999px;
            margin-top: 14px;
            font-size: 0.85rem;
            font-weight: 600;
        }

        .container {
            max-width: 1200px;
            margin: -25px auto 0;
            padding: 0 20px;
        }

        .controls {
            background: white;
            padding: 16px 24px;
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(0,0,0,0.06);
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 30px;
            border: 1px solid var(--border-color);
        }

        .search-box {
            flex: 1;
            min-width: 250px;
            position: relative;
        }

        .search-box input {
            width: 100%;
            padding: 12px 18px;
            border-radius: 12px;
            border: 1px solid var(--border-color);
            font-size: 0.95rem;
            font-family: inherit;
            outline: none;
            transition: border 0.2s;
        }

        .search-box input:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(13, 148, 136, 0.15);
        }

        .btn-print {
            background-color: var(--primary);
            color: white;
            border: none;
            padding: 12px 22px;
            border-radius: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .btn-print:hover {
            background-color: var(--primary-dark);
            transform: translateY(-1px);
        }

        .grid-columns {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
            gap: 24px;
        }

        .column-card {
            background: white;
            border-radius: 18px;
            border: 1px solid var(--border-color);
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.03);
            transition: all 0.25s ease;
            display: flex;
            flex-direction: column;
        }

        .column-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 30px rgba(0,0,0,0.08);
            border-color: #99f6e4;
        }

        .card-header {
            background: #f0fdfa;
            border-bottom: 2px solid #ccfbf1;
            padding: 16px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .card-header .num-tag {
            background: var(--primary);
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            width: 32px;
            height: 32px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .card-header .title-area {
            text-align: right;
            flex: 1;
            margin-left: 12px;
        }

        .card-header .title-arab {
            font-family: 'Amiri', serif;
            font-size: 1.55rem;
            color: var(--primary-dark);
            direction: rtl;
            font-weight: 700;
            line-height: 1.2;
        }

        .card-header .title-latin {
            font-size: 0.85rem;
            color: var(--text-muted);
            font-weight: 600;
        }

        .card-body {
            padding: 18px 20px;
            flex: 1;
        }

        .kaidah-box {
            background: #f8fafc;
            border-left: 3px solid var(--accent);
            padding: 8px 12px;
            border-radius: 0 8px 8px 0;
            margin-bottom: 14px;
            font-size: 0.82rem;
            color: #475569;
        }

        .word-list {
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }

        .word-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 8px 12px;
            background: #fafafa;
            border-radius: 10px;
            border: 1px solid #f1f5f9;
        }

        .word-item:hover {
            background: #f0fdf4;
        }

        .word-arab {
            font-family: 'Amiri', serif;
            font-size: 1.45rem;
            color: #047857;
            direction: rtl;
            font-weight: 700;
        }

        .word-info {
            flex: 1;
            margin: 0 14px;
            text-align: left;
        }

        .word-latin {
            font-size: 0.82rem;
            color: var(--text-muted);
            font-style: italic;
        }

        .word-arti {
            font-weight: 600;
            color: #1e293b;
            font-size: 0.95rem;
        }

        .freq-badge {
            background: #fee2e2;
            color: #b91c1c;
            font-size: 0.75rem;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 9999px;
            border: 1px solid #fca5a5;
            white-space: nowrap;
        }

        .freq-badge.empty {
            background: #f1f5f9;
            color: #94a3b8;
            border-color: #e2e8f0;
        }

        footer {
            margin-top: 50px;
            text-align: center;
            color: var(--text-muted);
            font-size: 0.9rem;
        }

        @media print {
            @page {
                size: A4;
                margin: 10mm;
            }
            body {
                background: white !important;
                padding: 0 !important;
                font-size: 9pt;
            }
            .header {
                background: none !important;
                color: #0f766e !important;
                padding: 5px 0 15px 0 !important;
                text-align: center !important;
                box-shadow: none !important;
            }
            .header h1 {
                font-size: 18pt !important;
                color: #0f766e !important;
            }
            .header p {
                font-size: 9.5pt !important;
                color: #334155 !important;
            }
            .header .badge-total {
                background: #f0fdfa !important;
                color: #0f766e !important;
                border: 1px solid #0f766e !important;
                font-size: 8pt !important;
                padding: 3px 10px !important;
            }
            .container {
                margin-top: 0 !important;
                padding: 0 !important;
                max-width: 100% !important;
            }
            .controls, .btn-print {
                display: none !important;
            }
            .grid-columns {
                display: grid !important;
                grid-template-columns: repeat(2, 1fr) !important;
                gap: 12px !important;
            }
            .column-card {
                break-inside: avoid !important;
                page-break-inside: avoid !important;
                box-shadow: none !important;
                border: 1px solid #cbd5e1 !important;
                border-radius: 12px !important;
            }
            .card-header {
                padding: 8px 12px !important;
            }
            .card-header .title-arab {
                font-size: 1.25rem !important;
            }
            .card-body {
                padding: 8px 12px !important;
            }
            .kaidah-box {
                font-size: 7.5pt !important;
                padding: 5px 8px !important;
                margin-bottom: 8px !important;
            }
            .word-item {
                padding: 4px 8px !important;
                border-radius: 6px !important;
            }
            .word-arab {
                font-size: 1.2rem !important;
            }
            .word-arti {
                font-size: 8.5pt !important;
            }
            .word-latin {
                font-size: 7.5pt !important;
            }
            .freq-badge {
                font-size: 7pt !important;
                padding: 2px 6px !important;
            }
            footer {
                margin-top: 20px !important;
                font-size: 8pt !important;
            }
        }
    </style>
</head>
<body>

    <header class="header">
        <h1>KOMPILASI 26 KOLOM METODE TAMYIZ</h1>
        <p>Kamus Rangkuman Huruf, Kata Hubung, Kata Tanya, dan Kata Ganti yang Paling Sering Berulang dalam Ayat-Ayat Suci Al-Qur'an</p>
        <div class="badge-total">Lengkap Kolom 1 s/d 26 Beserta Harakat, Arti & Frekuensi Mushaf</div>
    </header>

    <div class="container">
        <div class="controls">
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="Ketik lafadz Arab, arti, atau nama kolom untuk mencari..." onkeyup="filterCards()">
            </div>
            <button class="btn-print" onclick="window.print()">
                <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
                Cetak / Simpan PDF
            </button>
        </div>

        <div class="grid-columns" id="cardsContainer">
"""
    for col in COLUMNS_DATA:
        html_content += f"""
            <div class="column-card" data-search="{col['nomor']} {col['judul_arab']} {col['judul_latin']} {col['kategori']} {' '.join([x['arti'] + ' ' + x['latin'] + ' ' + x['arab'] for x in col['items']])}">
                <div class="card-header">
                    <div class="num-tag">{col['nomor']}</div>
                    <div class="title-area">
                        <div class="title-arab">{col['judul_arab']}</div>
                        <div class="title-latin">{col['judul_latin']}</div>
                    </div>
                </div>
                <div class="card-body">
                    <div class="kaidah-box">
                        <strong>Kaidah:</strong> {col['kaidah']}
                    </div>
                    <ul class="word-list">
        """
        for item in col['items']:
            freq_class = "freq-badge" if item['freq'] != "-" else "freq-badge empty"
            freq_text = f"{item['freq']}x" if item['freq'] != "-" else "-"
            html_content += f"""
                        <li class="word-item">
                            <span class="word-arab">{item['arab']}</span>
                            <div class="word-info">
                                <div class="word-arti">{item['arti']}</div>
                                <div class="word-latin">{item['latin']}</div>
                            </div>
                            <span class="{freq_class}" title="Frekuensi di Al-Qur'an">{freq_text}</span>
                        </li>
            """
        html_content += """
                    </ul>
                </div>
            </div>
        """

    html_content += """
        </div>

        <footer>
            <p><strong>Sumber Materi:</strong> Modul Pembelajaran Tamyiz Online (Metode Tamyiz - Terjemah Al-Qur'an)</p>
            <p>Disusun sebagai bahan ajar dan referensi belajar keluarga & pendidikan anak.</p>
        </footer>
    </div>

    <script>
        function filterCards() {
            const query = document.getElementById('searchInput').value.toLowerCase();
            const cards = document.querySelectorAll('.column-card');
            cards.forEach(card => {
                const searchData = card.getAttribute('data-search').toLowerCase();
                if (searchData.includes(query)) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });
        }
    </script>
</body>
</html>
"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] HTML saved: {filepath}")

def build_docx(filepath):
    doc = docx.Document()
    
    # Page Setup (A4)
    sec = doc.sections[0]
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.5)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.6)
    sec.right_margin = Inches(0.6)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # Helper XML functions
    def set_cell_background(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def set_cell_borders(cell, top="none", bottom="none", left="none", right="none", color="CCCCCC", sz="4"):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for border_name, border_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            if border_val != "none":
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), border_val)
                border.set(qn('w:sz'), sz)
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), color)
                tcBorders.append(border)
            else:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'none')
                tcBorders.append(border)
        tcPr.append(tcBorders)

    def add_arabic_run(paragraph, text, size_pt=14, bold=True, color_rgb=(0x04, 0x78, 0x57)):
        run = paragraph.add_run(text)
        run.bold = bold
        run.font.size = Pt(size_pt)
        run.font.color.rgb = RGBColor(*color_rgb)
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Traditional Arabic')
        rFonts.set(qn('w:cs'), 'Traditional Arabic')
        rFonts.set(qn('w:hAnsi'), 'Traditional Arabic')
        rPr.append(rFonts)
        rPr.append(parse_xml(f'<w:rtl {nsdecls("w")}/>'))
        return run

    # Cover / Header Banner
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = title_p.add_run("KOMPILASI LENGKAP 26 KOLOM METODE TAMYIZ\n")
    r1.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)

    r2 = title_p.add_run("Kamus Terjemah Huruf, Kata Hubung & Kata Ganti Al-Qur'an\n")
    r2.bold = True
    r2.font.size = Pt(12)
    r2.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    r3 = title_p.add_run("Dilengkapi Lafadz Arab Berharakat, Arti Bahasa Indonesia, Kaidah I'rab, dan Frekuensi di Al-Qur'an")
    r3.italic = True
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    # Info box
    info_table = doc.add_table(rows=1, cols=1)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False
    info_table.columns[0].width = Inches(7.07)
    c_info = info_table.cell(0, 0)
    set_cell_background(c_info, "F0FDFA")
    set_cell_margins(c_info, top=100, bottom=100, left=150, right=150)
    set_cell_borders(c_info, left="single", color="0F766E", sz="24")
    p_info = c_info.paragraphs[0]
    p_info.add_run("Petunjuk Belajar: ").bold = True
    p_info.add_run(
        "Metode Tamyiz menyederhanakan pemahaman Al-Qur'an dengan menghafal 26 Kolom Huruf/Kata Bantu ini. "
        "Kata-kata ini memiliki frekuensi pengulangan ribuan kali dalam Al-Qur'an. "
        "Angka pada kolom 'Frekuensi' menunjukkan perkiraan jumlah kemunculan kata tersebut di dalam Al-Qur'an."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Add each column as a beautifully formatted table
    for col in COLUMNS_DATA:
        # Header for the column
        hdr_table = doc.add_table(rows=1, cols=2)
        hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_table.autofit = False
        hdr_table.columns[0].width = Inches(4.5)
        hdr_table.columns[1].width = Inches(2.57)

        c_l = hdr_table.cell(0, 0)
        c_r = hdr_table.cell(0, 1)

        set_cell_background(c_l, "0F766E")
        set_cell_background(c_r, "0F766E")
        set_cell_margins(c_l, 90, 90, 140, 140)
        set_cell_margins(c_r, 90, 90, 140, 140)

        p_l = c_l.paragraphs[0]
        run_kol = p_l.add_run(f"KOLOM {col['nomor']} : {col['judul_latin'].upper()}\n")
        run_kol.bold = True
        run_kol.font.size = Pt(10)
        run_kol.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        run_kai = p_l.add_run(f"Kaidah: {col['kaidah']}")
        run_kai.italic = True
        run_kai.font.size = Pt(8.5)
        run_kai.font.color.rgb = RGBColor(0xCC, 0xFB, 0xF1)

        p_r = c_r.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_arabic_run(p_r, col['judul_arab'], size_pt=15, bold=True, color_rgb=(0xFF, 0xFF, 0xFF))

        # Table data
        num_items = len(col['items'])
        table = doc.add_table(rows=num_items + 1, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        col_widths = [Inches(0.5), Inches(1.8), Inches(1.6), Inches(2.17), Inches(1.0)]
        for i, w in enumerate(col_widths):
            table.columns[i].width = w

        # Header Row
        headers = ["No", "Lafadz Arab", "Transliterasi", "Terjemah Indonesia", "Frekuensi"]
        for j, h in enumerate(headers):
            cell = table.cell(0, j)
            set_cell_background(cell, "E2E8F0")
            set_cell_margins(cell, 60, 60, 80, 80)
            set_cell_borders(cell, bottom="single", color="94A3B8", sz="12")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in [0, 4] else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(h)
            r.bold = True
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

        # Data Rows
        for row_idx, it in enumerate(col['items'], 1):
            bg = "FFFFFF" if row_idx % 2 != 0 else "F8FAFC"
            
            # Col 0: No
            c0 = table.cell(row_idx, 0)
            set_cell_background(c0, bg)
            set_cell_margins(c0, 50, 50, 60, 60)
            set_cell_borders(c0, bottom="single", color="E2E8F0", sz="4")
            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r0 = p0.add_run(str(row_idx))
            r0.font.size = Pt(9)

            # Col 1: Arab
            c1 = table.cell(row_idx, 1)
            set_cell_background(c1, bg)
            set_cell_margins(c1, 50, 50, 80, 80)
            set_cell_borders(c1, bottom="single", color="E2E8F0", sz="4")
            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            add_arabic_run(p1, it['arab'], size_pt=13, bold=True, color_rgb=(0x04, 0x78, 0x57))

            # Col 2: Latin
            c2 = table.cell(row_idx, 2)
            set_cell_background(c2, bg)
            set_cell_margins(c2, 50, 50, 80, 80)
            set_cell_borders(c2, bottom="single", color="E2E8F0", sz="4")
            p2 = c2.paragraphs[0]
            r2 = p2.add_run(it['latin'])
            r2.italic = True
            r2.font.size = Pt(9)
            r2.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

            # Col 3: Arti
            c3 = table.cell(row_idx, 3)
            set_cell_background(c3, bg)
            set_cell_margins(c3, 50, 50, 80, 80)
            set_cell_borders(c3, bottom="single", color="E2E8F0", sz="4")
            p3 = c3.paragraphs[0]
            r3 = p3.add_run(it['arti'])
            r3.bold = True
            r3.font.size = Pt(9.5)
            r3.font.color.rgb = RGBColor(0x11, 0x18, 0x27)

            # Col 4: Freq
            c4 = table.cell(row_idx, 4)
            set_cell_background(c4, bg)
            set_cell_margins(c4, 50, 50, 60, 60)
            set_cell_borders(c4, bottom="single", color="E2E8F0", sz="4")
            p4 = c4.paragraphs[0]
            p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if it['freq'] != "-":
                r4 = p4.add_run(f"{it['freq']}x")
                r4.bold = True
                r4.font.size = Pt(8.5)
                r4.font.color.rgb = RGBColor(0xDC, 0x26, 0x26)
            else:
                r4 = p4.add_run("-")
                r4.font.size = Pt(8.5)
                r4.font.color.rgb = RGBColor(0x9C, 0xA3, 0xAF)

        # Spacer between columns
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(6)

    doc.save(filepath)
    print(f"[OK] DOCX saved: {filepath}")

if __name__ == "__main__":
    target_dir = r"D:\SYSTEM WINDOWS SUPPORT\PROGRAM\CLAUDE\PROJECT\Program Pendidikan Anak"
    md_file = os.path.join(target_dir, "Kompilasi_Metode_Tamyiz_26_Kolom.md")
    html_file = os.path.join(target_dir, "Kompilasi_Metode_Tamyiz_26_Kolom.html")
    docx_file = os.path.join(target_dir, "Kompilasi_Metode_Tamyiz_26_Kolom.docx")

    build_markdown(md_file)
    build_html(html_file)
    build_docx(docx_file)
    print("All documents generated successfully!")
