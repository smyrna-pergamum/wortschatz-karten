"""Wortschatz Karten tanıtım sitesini üç dilde üretir.

Kullanım:  python build.py
Girdi:     src/template.html, src/_defs.html, src/_mascot.html ve aşağıdaki S sözlüğü
Çıktı:     ../docs/ altında de/, en/, tr/ ve dil seçen kök index.html (GitHub Pages: main → /docs)

Bir metni değiştirmek için S içindeki ilgili anahtarı düzenleyip betiği yeniden çalıştır.
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).parent
OUT = ROOT.parent / "docs"
LANGS = ["de", "en", "tr"]
LANG_NAMES = {"de": "DE", "en": "EN", "tr": "TR"}
LANG_TITLES = {"de": "Deutsch", "en": "English", "tr": "Türkçe"}
XLSX = "https://github.com/smyrna-pergamum/wortschatz-karten/raw/main/DSH_Wortschatz_Table.xlsx"

# Her anahtar: {"de": ..., "en": ..., "tr": ...}. HTML etiketleri serbest.
S = {
    "title": {
        "de": "Wortschatz Karten — Es beginnt mit einem Wort!",
        "en": "Wortschatz Karten — It starts with a word!",
        "tr": "Wortschatz Karten — Bir kelimeyle başlar!",
    },
    "metaDesc": {
        "de": "Wortschatz Karten: kostenlose Karteikarten-App für Windows. Deutsch von A1 bis C1 und für die DSH. 1.436 Wörter, Leitner-System, Aussprache und Übersetzung in 10 Sprachen.",
        "en": "Wortschatz Karten: free flashcard app for Windows. German from A1 to C1 and for the DSH exam. 1,436 words, Leitner spaced repetition, pronunciation and translations in 10 languages.",
        "tr": "Wortschatz Karten: Windows için ücretsiz kelime kartı uygulaması. A1'den C1'e ve DSH'ye kadar Almanca. 1.436 kelime, Leitner tekrar sistemi, telaffuz ve 10 dilde çeviri.",
    },
    "skip": {"de": "Zum Inhalt", "en": "Skip to content", "tr": "İçeriğe geç"},
    "logoAria": {"de": "Wortschatz Karten, zum Seitenanfang", "en": "Wortschatz Karten, back to top", "tr": "Wortschatz Karten, başa dön"},
    "navAria": {"de": "Hauptmenü", "en": "Main menu", "tr": "Ana menü"},
    "langAria": {"de": "Sprache", "en": "Language", "tr": "Dil"},
    "navFeatures": {"de": "Funktionen", "en": "Features", "tr": "Özellikler"},
    "navTry": {"de": "Karte testen", "en": "Try a card", "tr": "Kartı dene"},
    "navHow": {"de": "So funktioniert’s", "en": "How it works", "tr": "Nasıl çalışır?"},
    "navGuide": {"de": "Anleitung", "en": "Guide", "tr": "Kılavuz"},
    "navExcel": {"de": "Excel-Liste", "en": "Excel list", "tr": "Excel listesi"},
    "navFaq": {"de": "Fragen", "en": "FAQ", "tr": "Sorular"},
    "dlShort": {"de": "Download", "en": "Download", "tr": "İndir"},
    "dlFree": {"de": "Kostenlos laden", "en": "Free download", "tr": "Ücretsiz indir"},

    "heroTag": {
        "de": "DEUTSCH VON A1 BIS C1 &amp; DSH",
        "en": "GERMAN FROM A1 TO C1 &amp; DSH",
        "tr": "A1'DEN C1'E VE DSH'YE ALMANCA",
    },
    "heroL1": {"de": "Lerne ein Wort.", "en": "Learn a word.", "tr": "Bir kelime öğren."},
    "heroL2": {
        "de": "<span>Entdecke</span> <span>eine</span> <span>Welt!</span>",
        "en": "<span>Open</span> <span>a</span> <span>world!</span>",
        "tr": "<span>Bir</span> <span>dünya</span> <span>aç!</span>",
    },
    "heroLead": {
        "de": "Ein Wort zu sehen ist eins, es wirklich zu lernen etwas anderes. Mit Erklärungen, Beispielsätzen und Wiederholung im richtigen Abstand baust du deinen deutschen Wortschatz Schritt für Schritt auf — vom ersten A1-Wort bis zur DSH.",
        "en": "Seeing a word is one thing; really learning it is another. With explanations, example sentences and spaced repetition you build your German vocabulary step by step — from your first A1 word all the way to the DSH.",
        "tr": "Bir kelimeyi görmek başka, onu gerçekten öğrenmek başka. Açıklamalar, örnek cümleler ve aralıklı tekrarla Almanca kelime dağarcığını adım adım geliştir — ilk A1 kelimenden DSH'ye kadar.",
    },
    "heroTry": {"de": "Erst eine Karte testen", "en": "Try a card first", "tr": "Önce bir kartı dene"},
    "heroExcel": {"de": "Excel-Wortliste", "en": "Excel word list", "tr": "Excel kelime listesi"},
    "heroGuide": {"de": "Anleitung", "en": "User guide", "tr": "Kullanım kılavuzu"},
    "heroSub": {
        "de": "Windows 10 / 11 · 64 Bit · MIT-Lizenz<br>Die neueste Version gibt es auf GitHub.",
        "en": "Windows 10 / 11 · 64-bit · MIT license<br>Latest version on GitHub.",
        "tr": "Windows 10 / 11 · 64 bit · MIT lisansı<br>En güncel sürüm GitHub üzerinden.",
    },
    "burstFree": {"de": "100%<br>GRATIS!", "en": "100%<br>FREE!", "tr": "100%<br>ÜCRETSİZ!"},

    "numbersAria": {"de": "Kurz zur App", "en": "The app at a glance", "tr": "Uygulama hakkında kısa bilgiler"},
    "n1": {"de": "1.436", "en": "1,436", "tr": "1.436"},
    "n1t": {"de": "fertige deutsche<br>Wörter in 10 Einheiten", "en": "ready-made German<br>words in 10 units", "tr": "10 ünitede hazır<br>Almanca kelime"},
    "n2t": {"de": "für alle Niveaus,<br>bis zur DSH", "en": "for every level,<br>up to the DSH", "tr": "her seviye için,<br>DSH'ye kadar"},
    "n3t": {"de": "Boxen im<br>Leitner-System", "en": "boxes in the<br>Leitner system", "tr": "kutulu Leitner<br>tekrar sistemi"},
    "n4t": {"de": "Sprachen für<br>Übersetzungen", "en": "languages for<br>translations", "tr": "dilde kelime<br>çevirisi"},

    "tvTag": {"de": "JETZT LIVE", "en": "NOW SHOWING", "tr": "ŞİMDİ YAYINDA"},
    "tvH": {"de": "Ein Blick in die App.", "en": "A look inside the app.", "tr": "Uygulamanın içinden."},
    "tvP": {
        "de": "Karteikarten, Lernmodi und dein Fortschritt — alles auf einem Bildschirm. Schalte um und sieh dir auch das dunkle Design an.",
        "en": "Flashcards, learning modes and your progress — all on one screen. Switch the channel to see the dark theme too.",
        "tr": "Kelime kartları, çalışma modları ve ilerlemen — tek ekranda. Kanalı değiştir, koyu temayı da gör.",
    },
    "zoomMainAria": {"de": "Screenshot der App vergrößern", "en": "Enlarge app screenshot", "tr": "Uygulama ekran görüntüsünü büyüt"},
    "zoomStatsAria": {"de": "Lernstatistik vergrößern", "en": "Enlarge learning statistics", "tr": "Öğrenme istatistikleri ekranını büyüt"},
    "zoom": {"de": "GRÖSSER!", "en": "ZOOM!", "tr": "BÜYÜT!"},
    "themeAria": {"de": "Design des Screenshots", "en": "Screenshot theme", "tr": "Ekran görüntüsü teması"},
    "light": {"de": "Hell", "en": "Light", "tr": "Açık"},
    "dark": {"de": "Dunkel", "en": "Dark", "tr": "Koyu"},
    "capLight": {
        "de": "Karteikarten, Lernmodi und dein Fortschritt — auf einem Bildschirm.",
        "en": "Flashcards, learning modes and your progress — on one screen.",
        "tr": "Kelime kartları, çalışma modları ve ilerlemen — tek ekranda.",
    },
    "capDark": {
        "de": "Die Rückseite: Erklärung, Übersetzung, Synonyme und Antonyme.",
        "en": "The back of a card: explanation, translation, synonyms and antonyms.",
        "tr": "Kartın arka yüzü: açıklama, çeviri, eş ve zıt anlamlılar.",
    },
    "altLight": {
        "de": "Wortschatz Karten im hellen Design: Karteikarte, Filter, Lernstatistik und Bedienknöpfe",
        "en": "Wortschatz Karten light theme: flashcard, filters, learning statistics and controls",
        "tr": "Wortschatz Karten açık tema: kelime kartı, filtreler, öğrenme istatistikleri ve kontrol düğmeleri",
    },
    "altDark": {
        "de": "Wortschatz Karten im dunklen Design: deutsche Worterklärung, Übersetzungen und Beispielsatz",
        "en": "Wortschatz Karten dark theme: German explanation, translations and example sentence",
        "tr": "Wortschatz Karten koyu tema: Almanca kelime açıklaması, çeviriler ve örnek cümle",
    },
    "altStats": {
        "de": "Lernstatistik: Kartenzahlen, Leitner-Boxen, Trefferquote und Ergebnisse pro Einheit",
        "en": "Learning statistics: card counts, Leitner boxes, hit rate and results per unit",
        "tr": "Öğrenme istatistikleri: kart sayıları, Leitner kutuları, başarı oranı ve ünite bazlı sonuçlar",
    },
    "dlgLight": {"de": "App · Helles Design", "en": "App · Light theme", "tr": "Uygulama · Açık tema"},
    "dlgDark": {"de": "App · Dunkles Design", "en": "App · Dark theme", "tr": "Uygulama · Koyu tema"},
    "dlgStats": {"de": "Lernstatistik", "en": "Learning statistics", "tr": "Öğrenme istatistikleri"},
    "close": {"de": "Schließen", "en": "Close", "tr": "Kapat"},

    "fTag": {"de": "LERNERLEBNIS", "en": "LEARNING EXPERIENCE", "tr": "ÖĞRENME DENEYİMİ"},
    "fH": {
        "de": "Einen Schritt <span class=\"wobble\" style=\"color:var(--red)\">weiter</span> als Auswendiglernen.",
        "en": "One step <span class=\"wobble\" style=\"color:var(--red)\">beyond</span> memorising.",
        "tr": "Ezberin bir adım <span class=\"wobble\" style=\"color:var(--red)\">ötesinde.</span>",
    },
    "f1H": {"de": "Nicht nur die Bedeutung, auch den Kontext lernen.", "en": "Learn the context, not just the meaning.", "tr": "Sadece anlamını değil, bağlamını öğren."},
    "f1P": {
        "de": "Auf jeder Rückseite: deutsche Erklärung, Synonyme und Antonyme, Übersetzungen und Beispielsätze. Kein Wort steht allein.",
        "en": "Every card back has a German explanation, synonyms and antonyms, translations and example sentences. No word stands alone.",
        "tr": "Her kartın arkasında Almanca açıklama, eş ve zıt anlamlılar, çeviriler ve örnek cümleler. Kelimeler tek başına kalmasın.",
    },
    "f1N": {"de": "ERKLÄRUNG · SYNONYM · ANTONYM · BEISPIEL", "en": "EXPLANATION · SYNONYM · ANTONYM · EXAMPLE", "tr": "AÇIKLAMA · SYNONYM · ANTONYM · BEISPIEL"},
    "f2H": {"de": "Zurück zu dem, was du brauchst.", "en": "Go back to what you need, not what you know.", "tr": "Bildiklerine değil, ihtiyacın olanlara dön."},
    "f2P": {
        "de": "Neue Wörter, fällige Wiederholungen, zuletzt falsch beantwortete oder schwierige Wörter: Wähle den Lernmodus, der heute passt.",
        "en": "New words, cards due for review, recent mistakes or difficult words: pick the learning mode that fits your day.",
        "tr": "Yeni, tekrar zamanı gelen, son yanlış yanıtlanan veya zorlandığın kelimeler. Öğrenme modunu o günkü ihtiyacına göre seç.",
    },
    "f2N": {"de": "WIEDERHOLUNG · FAVORITEN · SCHWIERIGE WÖRTER", "en": "SPACED REPETITION · FAVORITES · DIFFICULT WORDS", "tr": "ARALIKLI TEKRAR · FAVORİLER · ZOR KELİMELER"},
    "f3H": {"de": "Finde deinen Rhythmus — und deine Stimme.", "en": "Find your own rhythm — and your voice.", "tr": "Kendi ritmini, kendi sesini bul."},
    "f3P": {
        "de": "Hör dir die deutsche Aussprache an und sieh die Bedeutung in deiner Muttersprache. Helles oder dunkles Design, alles per Tastatur.",
        "en": "Listen to the German pronunciation and see the meaning in your mother tongue. Light or dark theme, all from the keyboard.",
        "tr": "Almanca telaffuzu dinle, ana dilindeki karşılığa bak. Açık veya koyu temayla, klavye kısayollarını kullanarak çalış.",
    },
    "f3N": {"de": "AUSSPRACHE · 10 SPRACHEN · 2 DESIGNS", "en": "PRONUNCIATION · 10 LANGUAGES · 2 THEMES", "tr": "SESLİ TELAFFUZ · 10 DİL · İKİ TEMA"},

    "dTag": {"de": "EIN KLEINER ANFANG", "en": "A SMALL START", "tr": "KÜÇÜK BİR BAŞLANGIÇ"},
    "dH": {"de": "Vorne ein Wort. Hinten eine Verbindung.", "en": "A word on the front. A connection on the back.", "tr": "Bir yüzünde kelime. Diğerinde bir bağlantı."},
    "dP": {
        "de": "Versuch dich an die Bedeutung zu erinnern, bevor du die Karte umdrehst. Dann knüpfst du die Verbindung mit Erklärung, Übersetzung und Beispiel.",
        "en": "Try to recall the meaning before you flip the card. Then build the connection with the explanation, translation and example.",
        "tr": "Kartı çevirmeden önce anlamını hatırlamayı dene. Sonra açıklamayla, çeviriyle ve örnekle bağlantıyı kur.",
    },
    "d1": {"de": "Erst selbst testen.", "en": "Test yourself first.", "tr": "Önce kendini sına."},
    "d2": {"de": "Die Rückseite entdecken.", "en": "Explore the back of the card.", "tr": "Kartın arkasını keşfet."},
    "d3": {"de": "In der App „gewusst“ oder „nicht gewusst“ wählen.", "en": "In the app, choose “known” or “not known”.", "tr": "Uygulamada “bildim” veya “bilemedim” de."},
    "d4": {"de": "Den Wiederholungsplan dem Leitner-System überlassen.", "en": "Leave the review plan to the Leitner system.", "tr": "Tekrar planını Leitner sistemine bırak."},
    "cardTag": {"de": "BEISPIELKARTE", "en": "SAMPLE CARD", "tr": "ÖRNEK KELİME KARTI"},
    "sideFront": {"de": "VORDERSEITE / DEUTSCH", "en": "FRONT / GERMAN", "tr": "ÖN YÜZ / DEUTSCH"},
    "sideBack": {"de": "RÜCKSEITE / BEDEUTUNG &amp; KONTEXT", "en": "BACK / MEANING &amp; CONTEXT", "tr": "ARKA YÜZ / ANLAM &amp; BAĞLAM"},
    "cardKind": {"de": "NOMEN · MASKULIN", "en": "NOUN · MASCULINE", "tr": "İSİM · MASKULIN"},
    "cardHint": {"de": "Was sagt dir dieses Wort?", "en": "What does this word tell you?", "tr": "Bu kelime sana ne anlatıyor?"},
    "cardDef": {"de": "Mundart, regionale Sprache", "en": "dialect, regional speech", "tr": "lehçe, ağız"},
    "cardWait": {"de": "Nimm dir einen Moment zum Erinnern …", "en": "Take a moment to remember …", "tr": "Hatırlamaya bir an ayır…"},
    "flip": {"de": "Karte umdrehen ↻", "en": "Flip card ↻", "tr": "Kartı çevir ↻"},
    "flipBack": {"de": "Zurück zur Vorderseite ↻", "en": "Back to the front ↻", "tr": "Ön yüze dön ↻"},
    "liveFront": {"de": "Vorderseite: der Dialekt.", "en": "Front: der Dialekt.", "tr": "Ön yüz: der Dialekt."},
    "liveBack": {
        "de": "Rückseite: Mundart, regionale Sprache. In Bayern hört man viele Dialekte. Synonym: Mundart, Antonym: Standardsprache.",
        "en": "Back: dialect, regional speech. In Bayern hört man viele Dialekte. Synonym: Mundart, antonym: Standardsprache.",
        "tr": "Arka yüz: lehçe, ağız. In Bayern hört man viele Dialekte. Synonym: Mundart, Antonym: Standardsprache.",
    },
    "demoNote": {
        "de": "Interaktives Beispiel. Dieser Test wirkt sich nicht auf deinen Lernfortschritt in der App aus.",
        "en": "Interactive demo. Trying it here does not affect your progress in the app.",
        "tr": "Etkileşimli tanıtım örneği. Buradaki deneme, uygulamadaki öğrenme ilerlemeni etkilemez.",
    },

    "mTag": {"de": "LEITNER-SYSTEM", "en": "LEITNER SYSTEM", "tr": "LEITNER SİSTEMİ"},
    "mH": {"de": "Jedes Wort kommt zur richtigen Zeit zurück.", "en": "Come back to every word at the right time.", "tr": "Her kelimeye, doğru zamanda yeniden dön."},
    "mP": {
        "de": "Richtig beantwortete Karten wandern in die nächste Box, der Abstand wird größer. Falsch beantwortete kommen zurück in Box 1 und am selben Tag noch einmal dran. Klick auf eine Box!",
        "en": "Cards you answer correctly move to the next box and the interval grows. Wrong answers send a card back to box 1, and it comes up again the same day. Click a box!",
        "tr": "Doğru yanıtladığın kartlar bir sonraki kutuya geçer, tekrar aralığı uzar. Yanlış yanıtladığın kartlar birinci kutuya döner ve aynı gün yeniden çalışılır. Bir kutuya tıkla!",
    },
    "mAria": {"de": "Leitner-Wiederholungsabstände", "en": "Leitner review intervals", "tr": "Leitner tekrar aralıkları"},
    "box": {"de": "Box", "en": "Box", "tr": "Kutu"},
    "boxCaps": {"de": "BOX", "en": "BOX", "tr": "KUTU"},
    "days": {"de": ["Tag später", "Tage später"], "en": ["day later", "days later"], "tr": ["gün sonra", "gün sonra"]},
    "boxes": {
        "de": [["Erster Schritt", "Nach einer richtigen Antwort kommt die erste Wiederholung nach 1 Tag. Falsch beantwortete Karten kommen heute noch einmal."],
               ["Schon vertrauter", "In Box 2 beträgt der neue Abstand für eine richtig beantwortete Karte 2 Tage."],
               ["Die Verbindung wächst", "In Box 3 steigt der Abstand auf 4 Tage. Wörter, die du kennst, lernst du seltener."],
               ["Längerer Abstand", "Karten in Box 4 werden nach 8 Tagen wiederholt. Eine falsche Antwort heißt: zurück in Box 1."],
               ["Gelernte Wörter", "Karten in Box 5 gelten als gelernt und werden alle 16 Tage wiederholt."]],
        "en": [["First step", "After a correct answer the first review comes 1 day later. Cards you got wrong come back today."],
               ["A bit more familiar", "In box 2, a correctly answered card gets a new interval of 2 days."],
               ["The connection grows", "In box 3 the interval rises to 4 days. You study the words you remember less often."],
               ["Longer interval", "Cards in box 4 are reviewed after 8 days. A wrong answer means back to box 1."],
               ["Learned words", "Cards in box 5 count as learned and are reviewed every 16 days."]],
        "tr": [["İlk adım", "Doğru yanıtın ardından ilk tekrar 1 gün sonra. Yanlış yanıtlanan kartlar ise bugün yeniden karşına çıkar."],
               ["Biraz daha tanıdık", "İkinci kutuda doğru yanıtlanan kartın yeni tekrar aralığı 2 gündür."],
               ["Bağlantı güçleniyor", "Üçüncü kutuda yeni tekrar aralığı 4 güne çıkar. Hatırladığın kelimeleri daha seyrek çalışırsın."],
               ["Daha uzun aralık", "Dördüncü kutudaki kartlar 8 gün sonra tekrar edilir. Yanlış yanıt birinci kutuya dönüş demektir."],
               ["Öğrenilen kelimeler", "Beşinci kutudaki kartlar öğrenilmiş olarak sayılır ve 16 günlük aralıklarla tekrar edilir."]],
    },

    "gTag": {"de": "ANLEITUNG", "en": "USER GUIDE", "tr": "KULLANIM KILAVUZU"},
    "gH": {
        "de": "In fünf Schritten <span class=\"wobble\" style=\"color:var(--sky)\">zur ersten Karte!</span>",
        "en": "Five steps <span class=\"wobble\" style=\"color:var(--sky)\">to your first card!</span>",
        "tr": "Beş adımda <span class=\"wobble\" style=\"color:var(--sky)\">ilk kartına!</span>",
    },
    "gP": {
        "de": "Alles, was du von der Installation bis zur ersten Wiederholungsrunde brauchst. Die App ist auf Deutsch — die Menünamen stehen hier im Original.",
        "en": "Everything you need from installing to your first review round. The app’s interface is in German, so menu names are given in German below.",
        "tr": "Kurulumdan ilk tekrar turuna kadar bilmen gereken her şey. Uygulamanın arayüzü Almancadır; menü adlarını aşağıda Almanca olarak bulursun.",
    },
    "s1H": {"de": "Herunterladen", "en": "Download", "tr": "İndir"},
    "s1P": {
        "de": "Lade <code>WortschatzKarten_Setup.exe</code> aus der neuesten Version auf GitHub herunter.",
        "en": "Download <code>WortschatzKarten_Setup.exe</code> from the latest release on GitHub.",
        "tr": "GitHub’daki en güncel sürümden <code>WortschatzKarten_Setup.exe</code> dosyasını indir.",
    },
    "s1L": {"de": "Zur neuesten Version", "en": "Go to latest release", "tr": "Son sürüme git"},
    "s2H": {"de": "Installieren", "en": "Install", "tr": "Kur"},
    "s2P": {
        "de": "Starte das Setup. Die App wird unter <code>Program Files</code> installiert (Administratorrechte nötig), mit Startmenü-Eintrag und Deinstallationsprogramm.",
        "en": "Run the installer. It installs to <code>Program Files</code> (administrator rights required) and adds a Start menu entry and an uninstaller.",
        "tr": "Kurulum dosyasını çalıştır. Uygulama <code>Program Files</code> klasörüne kurulur (yönetici yetkisi gerekir); Başlat menüsü kısayolu ve kaldırıcı eklenir.",
    },
    "s3H": {"de": "Liste importieren", "en": "Import the list", "tr": "Listeyi içe aktar"},
    "s3P": {
        "de": "Beim ersten Start findet die App die mitgelieferte Wortliste und fragt, ob sie importiert werden soll. Klick auf <b>Ja</b> — 1.436 Wörter sind bereit.",
        "en": "On first start the app finds the included word list and asks whether to import it. Click <b>Yes</b> — 1,436 words are ready.",
        "tr": "İlk açılışta uygulama hazır kelime listesini bulur ve içe aktarmak isteyip istemediğini sorar. <b>Evet</b> de, 1.436 kelime hazır.",
    },
    "s4H": {"de": "Karte umdrehen", "en": "Flip the card", "tr": "Kartı çevir"},
    "s4P": {
        "de": "Sieh dir das Wort an, versuche dich an die Bedeutung zu erinnern, und dreh die Karte mit <kbd>Space</kbd> um.",
        "en": "Look at the word, try to recall its meaning, then press <kbd>Space</kbd> to flip the card.",
        "tr": "Kelimeyi gör, anlamını hatırlamaya çalış, sonra <kbd>Space</kbd> ile kartı çevir ve arka yüzü incele.",
    },
    "s5H": {"de": "Antworten", "en": "Answer", "tr": "Cevabını ver"},
    "s5P": {
        "de": "<kbd>2</kbd> „gewusst“, <kbd>1</kbd> „nicht gewusst“. Das Leitner-System legt die Karte in die richtige Box und bringt sie zur richtigen Zeit zurück.",
        "en": "<kbd>2</kbd> “known”, <kbd>1</kbd> “not known”. The Leitner system puts the card in the right box and brings it back at the right time.",
        "tr": "<kbd>2</kbd> “bildim”, <kbd>1</kbd> “bilemedim”. Kart, Leitner sistemine göre doğru kutuya geçer ve doğru zamanda geri gelir.",
    },
    "smart": {
        "de": "<b>Falls Windows warnt:</b> Das Setup ist noch nicht signiert. SmartScreen kann vor einem „unbekannten Herausgeber“ warnen: <b>Weitere Informationen → Trotzdem ausführen</b>. Du kannst den Download mit der SHA-256-Prüfsumme in den Release-Notes vergleichen.",
        "en": "<b>If Windows warns you:</b> the installer is not code-signed yet, so SmartScreen may warn about an unknown publisher: <b>More info → Run anyway</b>. You can check the download against the SHA-256 checksum in the release notes.",
        "tr": "<b>Windows uyarı verirse:</b> Kurulum dosyası henüz kod imzalı değil. SmartScreen “bilinmeyen yayımcı” uyarısı gösterebilir: <b>Ek bilgi → Yine de çalıştır</b>. İndirdiğin dosyayı sürüm notlarındaki SHA-256 sağlama değeriyle karşılaştırabilirsin.",
    },
    "kTag": {"de": "TASTENKÜRZEL", "en": "KEYBOARD SHORTCUTS", "tr": "KLAVYE KISAYOLLARI"},
    "kH": {"de": "Ganz ohne Maus.", "en": "No mouse needed.", "tr": "Fareye gerek yok."},
    "k1": {"de": "Karte umdrehen", "en": "Flip card", "tr": "Kartı çevir"},
    "k2": {"de": "Nächste Karte", "en": "Next card", "tr": "Sonraki kart"},
    "k3": {"de": "Vorherige Karte", "en": "Previous card", "tr": "Önceki kart"},
    "k4": {"de": "Nicht gewusst / Gewusst", "en": "Not known / Known", "tr": "Bilemedim / Bildim"},
    "k5": {"de": "Aussprache", "en": "Pronunciation", "tr": "Telaffuzu dinle"},
    "k6": {"de": "Favorit / Schwierig", "en": "Favorite / Difficult", "tr": "Favori / Zor kelime"},
    "k7": {"de": "Mischen", "en": "Shuffle", "tr": "Kartları karıştır"},
    "k8": {"de": "Excel-Datei aktualisieren", "en": "Update the Excel file", "tr": "Excel dosyasını güncelle"},
    "k9": {"de": "Favoriten &amp; schwierige Wörter exportieren", "en": "Export favorites &amp; difficult words", "tr": "Favori ve zor kelimeleri dışa aktar"},
    "modesTag": {"de": "LERNMODI", "en": "LEARNING MODES", "tr": "ÖĞRENME MODLARI"},
    "modesH": {"de": "Was brauchst du heute?", "en": "What do you need today?", "tr": "Bugün neye ihtiyacın var?"},
    "modesList": {
        "de": ["<b>Wiederholung:</b> fällige und neue Karten", "<b>Alle Karten:</b> die ganze Liste", "<b>Nur neue:</b> Wörter, die du noch nie gesehen hast", "<b>Zuletzt falsch:</b> deine letzten Fehler", "<b>Favoriten</b> und <b>schwierige Wörter</b>"],
        "en": ["<b>Review:</b> due cards plus new words", "<b>All cards:</b> the whole list", "<b>New only:</b> words you have never seen", "<b>Last wrong:</b> your most recent mistakes", "<b>Favorites</b> and <b>difficult words</b>"],
        "tr": ["<b>Tekrar:</b> zamanı gelen kartlar ve yeni kelimeler", "<b>Tüm kartlar:</b> listenin tamamı", "<b>Sadece yeniler:</b> hiç görmediğin kelimeler", "<b>Son yanlışlar:</b> en son bilemediklerin", "<b>Favoriler</b> ve <b>zor kelimeler</b>"],
    },
    "modesNote": {
        "de": "Alles lässt sich nach Einheit und Abschnitt filtern. Das Statistikfenster zeigt Leitner-Boxen, Trefferquote und Ergebnisse pro Einheit.",
        "en": "Filter any mode by unit and section. The statistics window shows the Leitner boxes, your hit rate and results per unit.",
        "tr": "Hepsini ünite ve bölüme göre filtreleyebilirsin. İstatistik penceresi Leitner kutularını, başarı oranını ve ünite sonuçlarını gösterir.",
    },
    "setTag": {"de": "EINSTELLUNGEN &amp; DATEN", "en": "SETTINGS &amp; DATA", "tr": "AYARLAR &amp; VERİLER"},
    "setH": {"de": "Alles auf deinem Rechner.", "en": "Everything stays on your PC.", "tr": "Her şey senin bilgisayarında."},
    "setList": {
        "de": ["Helles / dunkles Design und Muttersprache direkt in der App.", "Wortliste, Fortschritt und Einstellungen liegen in <code>%APPDATA%\\WortschatzKarten\\</code>.", "Neu anfangen: <b lang=\"de\">Daten → Lernfortschritt zurücksetzen</b>.", "Eigene Liste: <b lang=\"de\">Daten → Excel aktualisieren…</b> (<kbd>Ctrl</kbd>+<kbd>O</kbd>)."],
        "en": ["Light / dark theme and your mother tongue right in the app.", "Your word list, progress and settings live in <code>%APPDATA%\\WortschatzKarten\\</code>.", "Start over: <b lang=\"de\">Daten → Lernfortschritt zurücksetzen</b>.", "Your own list: <b lang=\"de\">Daten → Excel aktualisieren…</b> (<kbd>Ctrl</kbd>+<kbd>O</kbd>)."],
        "tr": ["Açık / koyu tema ve ana dil seçimi uygulama içinden.", "Listen, ilerlemen ve ayarların <code>%APPDATA%\\WortschatzKarten\\</code> klasöründe saklanır.", "Baştan başlamak için: <b lang=\"de\">Daten → Lernfortschritt zurücksetzen</b>.", "Kendi listen için: <b lang=\"de\">Daten → Excel aktualisieren…</b> (<kbd>Ctrl</kbd>+<kbd>O</kbd>)."],
    },

    "sheetBurst": {"de": "1.436<br>WÖRTER", "en": "1,436<br>WORDS", "tr": "1.436<br>KELİME"},
    "xTag": {"de": "EXCEL-WORTLISTE", "en": "EXCEL WORD LIST", "tr": "EXCEL KELİME LİSTESİ"},
    "xH": {"de": "Liste laden und nach Belieben anpassen.", "en": "Download the list, make it your own.", "tr": "Listeyi indir, istediğin gibi değiştir."},
    "xP": {
        "de": "Die mitgelieferte DSH-Wortliste ist eine einzige Excel-Datei. Du kannst sie ausdrucken, eigene Notizen ergänzen oder als Vorlage für deine eigene Liste nutzen — egal ob A1 oder C1.",
        "en": "The included DSH word list is a single Excel file. Print it, add your own notes, or use it as a template for your own list — whether you are at A1 or C1.",
        "tr": "Uygulamadaki hazır DSH kelime listesi tek bir Excel dosyası. Yazdırabilir, kendi notlarını ekleyebilir ya da kendi listen için şablon olarak kullanabilirsin — A1'de de olsan C1'de de.",
    },
    "xBtn": {"de": "Excel-Liste laden", "en": "Download Excel list", "tr": "Excel listesini indir"},
    "xView": {"de": "Auf GitHub ansehen", "en": "View on GitHub", "tr": "GitHub’da görüntüle"},
    "xColsS": {"de": "Spalten für deine eigene Liste", "en": "Columns for your own list", "tr": "Kendi listen için sütunlar"},
    "xColsP": {
        "de": "Die Kopfzeile muss in den ersten 15 Zeilen stehen. Deutsche, englische und türkische Spaltennamen werden erkannt. Nur die Spalte <b>Wort</b> ist Pflicht.",
        "en": "The header row must be within the first 15 rows. German, English and Turkish column names are recognised. Only the <b>word</b> column is required.",
        "tr": "Excel dosyanda başlık satırı ilk 15 satırın içinde olmalı. Almanca, İngilizce ve Türkçe sütun adları tanınır. Yalnızca <b>kelime</b> sütunu zorunludur.",
    },
    "cCol": {"de": "Spalte", "en": "Column", "tr": "Sütun"},
    "cMean": {"de": "Bedeutung", "en": "Meaning", "tr": "Anlamı"},
    "cReq": {"de": "Pflicht", "en": "Required", "tr": "Zorunlu"},
    "cUnit": {"de": "Einheit", "en": "Unit", "tr": "Ünite"},
    "cSection": {"de": "Abschnitt / Thema", "en": "Section / topic", "tr": "Bölüm / konu"},
    "cWord": {"de": "Deutsches Wort", "en": "German word", "tr": "Almanca kelime"},
    "cExpl": {"de": "Deutsche Erklärung", "en": "German explanation", "tr": "Almanca açıklama"},
    "cSyn": {"de": "Synonyme / Antonyme", "en": "Synonyms / antonyms", "tr": "Eş / zıt anlamlılar"},
    "cEn": {"de": "Englische Bedeutung", "en": "English meaning", "tr": "İngilizce anlam"},
    "cTr": {"de": "Türkische Bedeutung", "en": "Turkish meaning", "tr": "Türkçe anlam"},
    "cEx": {"de": "Beispielsatz", "en": "Example sentence", "tr": "Örnek cümle"},
    "yes": {"de": "ja", "en": "yes", "tr": "evet"},
    "no": {"de": "nein", "en": "no", "tr": "hayır"},
    "xColsNote": {
        "de": "* Ohne Einheiten-Spalte wird der Tabellenblattname verwendet. Beim erneuten Import wird dein Fortschritt über Einheit und Wort zugeordnet und bleibt erhalten.",
        "en": "* Without a unit column, the sheet name is used. When you re-import an updated list, progress is matched by unit and word and kept.",
        "tr": "* Ünite sütunu yoksa sayfa adı kullanılır. Listeyi güncelleyip yeniden içe aktardığında ilerlemen ünite ve kelimeye göre eşleştirilip korunur.",
    },

    "iTag": {"de": "FORTSCHRITT SEHEN", "en": "SEE YOUR PROGRESS", "tr": "İLERLEMENİ GÖR"},
    "iH": {"de": "Kleine Schritte, sichtbarer Fortschritt.", "en": "Small steps, visible progress.", "tr": "Küçük adımlar, görünür ilerleme."},
    "iP": {
        "de": "Wie viele Wörter hast du gelernt, wo hattest du Mühe, welche Karten warten auf dich? Die Statistik gibt dir den Überblick.",
        "en": "How many words have you studied, which unit was hard, which cards are waiting? The statistics window gives you the overview.",
        "tr": "Kaç kelime çalıştın, hangi ünitede zorlandın, hangi kartlar tekrar bekliyor? İstatistik ekranı çalışma sürecine toplu bir bakış sunar.",
    },
    "iCap": {
        "de": "Eine Beispielsitzung: Ergebnisse pro Einheit und Verteilung auf die Leitner-Boxen.",
        "en": "A sample session: results per unit and the distribution across Leitner boxes.",
        "tr": "Örnek bir çalışma oturumu: ünite bazında sonuçlar ve Leitner kutularının dağılımı.",
    },
    "asideAria": {"de": "Weitere Funktionen", "en": "More features", "tr": "Diğer uygulama özellikleri"},
    "a1Tag": {"de": "DEINE LISTE, DEIN ZIEL", "en": "YOUR LIST, YOUR GOAL", "tr": "KENDİ LİSTEN, KENDİ HEDEFİN"},
    "a1H": {"de": "Bring dein Vokabelheft mit.", "en": "Bring your own vocabulary book.", "tr": "Kelime defterini yanında getir."},
    "a1P": {
        "de": "Importiere deine eigene Excel-Datei — für jedes Niveau von A1 bis C1. Beim Aktualisieren bleibt dein Fortschritt erhalten. Favoriten und schwierige Wörter kannst du nach Excel exportieren.",
        "en": "Import your own Excel file — for any level from A1 to C1. Updating it keeps your progress. Export favorites and difficult words to Excel.",
        "tr": "Kendi Excel dosyanı içe aktar — A1'den C1'e her seviye için. Listeyi güncellediğinde ilerlemen korunur. Favorilerini ve zorlandığın kelimeleri Excel’e aktarabilirsin.",
    },
    "a2Tag": {"de": "BRÜCKE ZUR MUTTERSPRACHE", "en": "A BRIDGE TO YOUR LANGUAGE", "tr": "ANA DİLİNDE BİR KÖPRÜ"},
    "a2H": {"de": "Ein Wort. In vielen Sprachen.", "en": "One word. Many languages.", "tr": "Bir kelime. Farklı dillerde karşılığı."},
    "a2P": {
        "de": "Sieh das deutsche Wort und den Beispielsatz in einer der unterstützten Sprachen.",
        "en": "See the German word and example sentence in one of the supported languages.",
        "tr": "Almanca kelimeyi ve örnek cümleyi desteklenen dillerden birinde görüntüle.",
    },
    "langList": {
        "de": ["Türkisch", "Englisch", "Arabisch", "Russisch", "Persisch", "Spanisch", "Französisch", "Italienisch", "Polnisch", "Ukrainisch"],
        "en": ["Turkish", "English", "Arabic", "Russian", "Persian", "Spanish", "French", "Italian", "Polish", "Ukrainian"],
        "tr": ["Türkçe", "İngilizce", "Arapça", "Rusça", "Farsça", "İspanyolca", "Fransızca", "İtalyanca", "Lehçe", "Ukraynaca"],
    },
    "a3Tag": {"de": "OHNE KONTO", "en": "NO ACCOUNT", "tr": "HESAPSIZ BAŞLA"},
    "a3H": {"de": "Dein Lernraum, auf deinem Rechner.", "en": "Your study space, on your own PC.", "tr": "Öğrenme alanın, kendi bilgisayarında."},
    "a3P": {
        "de": "Kein Konto nötig. Wortlisten, Fortschritt und Einstellungen bleiben auf deinem Rechner. Details zu optionaler Online-Übersetzung und Aussprache findest du unten.",
        "en": "No account needed. Word lists, progress and settings stay on your computer. Details about the optional online translation and pronunciation are below.",
        "tr": "Hesap açman gerekmez. Kelime listelerin, ilerlemen ve ayarların bilgisayarında saklanır. İsteğe bağlı çevrimiçi çeviri ve telaffuzun ayrıntıları aşağıda.",
    },

    "qTag": {"de": "HÄUFIGE FRAGEN", "en": "FREQUENTLY ASKED", "tr": "MERAK EDİLENLER"},
    "qH": {"de": "Bevor du loslegst.", "en": "Before you start.", "tr": "Başlamadan önce."},
    "qIntro": {"de": "Eine eher technische Frage?", "en": "A more technical question?", "tr": "Daha teknik bir soru mu var?"},
    "qDocs": {"de": "Zur Projektdokumentation", "en": "Read the project docs", "tr": "Proje belgelerine göz at"},
    "faq": {
        "de": [
            ("Für wen ist die App?", "Für alle, die Deutsch lernen — von A1 bis C1 — und für alle, die sich auf die DSH vorbereiten. Die mitgelieferte Liste enthält den DSH-Wortschatz; mit deiner eigenen Excel-Liste lernst du auf jedem Niveau. Die Oberfläche der App ist auf Deutsch."),
            ("Ist sie wirklich kostenlos?", "Ja. Wortschatz Karten ist kostenlos und Open Source unter der MIT-Lizenz. Den Quellcode findest du auf GitHub. Ein Konto brauchst du nicht."),
            ("Auf welchen Geräten läuft sie?", "Das Setup ist für Windows 10 und Windows 11 (64 Bit). Die Karte auf dieser Seite ist nur ein Beispiel; die vollständige App wird unter Windows installiert."),
            ("Brauche ich Internet?", "Karteikarten und Lernfortschritt funktionieren offline. Nur die optionale Übersetzung und die Online-Aussprache brauchen Internet; Ergebnisse werden zwischengespeichert. Für die Übersetzung werden das deutsche Wort und der Beispielsatz an Google Translate gesendet (Sätze notfalls an MyMemory), für die Online-Aussprache das Wort an Googles Sprachdienst. Mit einer deutschen Windows-Stimme klappt die Aussprache auch offline."),
            ("Kann ich eigene Wörter hinzufügen?", f"Ja. Über <span lang=\"de\">Daten → Excel aktualisieren…</span> importierst du deine Excel-Datei (.xlsx). Nur die Wortspalte ist Pflicht; Erklärung, Übersetzung, Synonyme, Antonyme und Beispielsätze sind optional. <a href=\"{XLSX}\" download>Beispieldatei herunterladen.</a> Die Spalten findest du im <a href=\"#excel\">Excel-Abschnitt</a>."),
            ("Wie installiere ich die App?", "Der Download-Knopf führt zur neuesten Version auf GitHub. Lade dort <strong>WortschatzKarten_Setup.exe</strong> herunter und folge dem Setup (Administratorrechte nötig). Bestätige beim ersten Start den Import der Wortliste. Schritt für Schritt: <a href=\"#kilavuz\">Anleitung</a>."),
        ],
        "en": [
            ("Who is the app for?", "For anyone learning German — from A1 to C1 — and for students preparing for the DSH. The included list covers DSH vocabulary; with your own Excel list you can study at any level. The app’s interface is in German."),
            ("Is it really free?", "Yes. Wortschatz Karten is free and open source under the MIT license. The source code is on GitHub. No account is needed."),
            ("Which devices does it run on?", "The installer is for Windows 10 and Windows 11, 64-bit. The card on this page is only a demo; the full app is installed on Windows."),
            ("Do I need an internet connection?", "Flashcards and progress tracking work offline. The optional translation and online pronunciation need internet; results are cached. For translation, the German word and example sentence are sent to Google Translate (sentences fall back to MyMemory); for online pronunciation, the word is sent to Google’s text-to-speech service. With a German Windows voice, pronunciation also works offline."),
            ("Can I add my own words?", f"Yes. Use <span lang=\"de\">Daten → Excel aktualisieren…</span> to import your Excel (.xlsx) file. Only the word column is required; explanation, translation, synonyms, antonyms and example sentences are optional. <a href=\"{XLSX}\" download>Download the sample file.</a> The column list is in the <a href=\"#excel\">Excel section</a>."),
            ("How do I install it?", "The download button takes you to the latest release on GitHub. Download <strong>WortschatzKarten_Setup.exe</strong> and follow the installer (administrator rights required). On first start, accept the offer to import the word list. Step by step: <a href=\"#kilavuz\">user guide</a>."),
        ],
        "tr": [
            ("Uygulama kimler için?", "A1'den C1'e Almanca öğrenen herkes ve DSH sınavına hazırlananlar için. Hazır liste DSH kelimelerini içerir; kendi Excel listenle her seviyede çalışabilirsin. Uygulamanın arayüz dili Almancadır."),
            ("Gerçekten ücretsiz mi?", "Evet. Wortschatz Karten ücretsiz ve MIT lisansıyla açık kaynak olarak sunulur. Kaynak kodunu GitHub üzerinden inceleyebilirsin. Hesap açılması gerekmez."),
            ("Hangi cihazlarda çalışır?", "Hazır kurulum dosyası Windows 10 ve Windows 11, 64 bit sistemler içindir. Bu sayfadaki kart bir tanıtım örneğidir; tam uygulama Windows’a kurulur."),
            ("İnternet bağlantısı gerekiyor mu?", "Kelime kartları ve öğrenme takibi çevrimdışı çalışır. İsteğe bağlı çeviri ve çevrimiçi telaffuz için internet gerekir; sonuçlar önbelleğe alınır. Çeviri için Almanca kelime ve örnek cümle Google Translate’e, cümle çevirisinde gerektiğinde MyMemory’ye; çevrimiçi telaffuz için kelime Google’ın ses hizmetine gönderilir. Uygun Almanca Windows sesi varsa çevrimdışı telaffuz da kullanılabilir."),
            ("Kendi kelimelerimi ekleyebilir miyim?", f"Evet. Uygulamadaki <span lang=\"de\">Daten → Excel aktualisieren…</span> menüsüyle Excel (.xlsx) dosyanı içe aktarabilirsin. Kelime sütunu zorunludur; açıklama, çeviri, eş ve zıt anlamlılar ile örnek cümleler isteğe bağlıdır. <a href=\"{XLSX}\" download>Örnek Excel dosyasını indir.</a> Sütunların listesi <a href=\"#excel\">Excel bölümünde</a>."),
            ("Nasıl yükleyebilirim?", "İndirme düğmesi seni GitHub’daki en güncel sürüme götürür. <strong>WortschatzKarten_Setup.exe</strong> dosyasını indirip kurulum adımlarını izle (yönetici yetkisi gerekir). İlk açılışta hazır kelime listesini içe aktarma önerisini onayla. Adım adım: <a href=\"#kilavuz\">kullanım kılavuzu</a>."),
        ],
    },

    "zTag": {"de": "BEREIT FÜR DEIN NÄCHSTES WORT?", "en": "READY FOR YOUR NEXT WORD?", "tr": "SONRAKİ KELİMENE HAZIR MISIN?"},
    "zH": {"de": "Starte heute mit einer Karte!", "en": "Start today with one card!", "tr": "Bugün bir kartla başla!"},
    "zP": {
        "de": "Mit der fertigen DSH-Liste oder deiner eigenen. Ein bisschen Lernen, ein neuer Anfang.",
        "en": "With the ready-made DSH list or your own. A little practice, a new beginning.",
        "tr": "Hazır DSH kelimeleriyle ya da kendi listenle. Küçük bir çalışma, yeni bir başlangıç.",
    },
    "zSource": {"de": "Quellcode", "en": "Source code", "tr": "Kaynak kodu"},
    "pr1": {"de": "KOSTENLOS &amp; OPEN SOURCE", "en": "FREE &amp; OPEN SOURCE", "tr": "ÜCRETSİZ &amp; AÇIK KAYNAK"},
    "pr2": {"de": "KEIN KONTO NÖTIG", "en": "NO ACCOUNT NEEDED", "tr": "HESAP GEREKTİRMEZ"},
    "pr3": {"de": "FORTSCHRITT BLEIBT BEI DIR", "en": "PROGRESS STAYS ON YOUR PC", "tr": "İLERLEME BİLGİSAYARINDA KALIR"},
    "feedback": {"de": "Feedback", "en": "Feedback", "tr": "Geri bildirim"},
    "license": {"de": "MIT-Lizenz", "en": "MIT license", "tr": "MIT lisansı"},
    "top": {"de": "Nach oben", "en": "Back to top", "tr": "Başa dön"},
    "footLine": {"de": "Es beginnt mit einem Wort. Jede Wiederholung ist ein Schritt weiter!", "en": "It starts with a word. Every review is one step further!", "tr": "Bir kelimeyle başlar. Her tekrar bir sonraki adıma!"},
    "footSmall": {"de": "Ein unabhängiges Open-Source-Werkzeug zum Deutschlernen.", "en": "An independent, open-source tool for learning German.", "tr": "Bağımsız, açık kaynak bir Almanca öğrenme aracı."},
}

# JavaScript'in ihtiyaç duyduğu anahtarlar
JS_KEYS = ["altLight", "altDark", "altStats", "capLight", "capDark", "sideFront", "sideBack", "flip", "flipBack",
           "liveFront", "liveBack", "box", "boxes", "dlgLight", "dlgDark", "dlgStats"]
DAYS = [1, 2, 4, 8, 16]


def crates(lang):
    out = []
    caps, days, boxes = S["boxCaps"][lang], S["days"][lang], S["boxes"][lang]
    for i, d in enumerate(DAYS):
        fly = "★" if i == 4 else "ABCD"[i]
        unit = days[0] if d == 1 else days[1]
        out.append(
            f'<button type="button" class="crate" data-box="{i+1}" aria-pressed="{"true" if i == 0 else "false"}">'
            f'<span class="flycard" aria-hidden="true">{fly}</span><span class="box"><span class="box-label">{caps} {i+1}</span>'
            f'<span class="box-time">{d}<small>{unit}</small></span><span class="box-level">{boxes[i][0]}</span></span></button>')
    return "\n      ".join(out)


def unescape_js(v):
    # JS textContent ile yazılan metinlerde HTML varlıkları olmasın
    if isinstance(v, str):
        return v.replace("&amp;", "&")
    if isinstance(v, list):
        return [unescape_js(x) for x in v]
    return v


def page(lang, tpl, partials):
    v = {k: val[lang] for k, val in S.items() if isinstance(val, dict) and lang in val}
    v["lang"] = lang
    v["langSwitch"] = "".join(
        f'<a href="../{l}/" hreflang="{l}" lang="{l}" title="{LANG_TITLES[l]}"' + (' aria-current="page"' if l == lang else "") + f'>{LANG_NAMES[l]}</a>'
        for l in LANGS)
    v["crates"] = crates(lang)
    v["box1Title"] = f'{S["box"][lang]} 1 · {S["boxes"][lang][0][0]}'
    v["box1Text"] = S["boxes"][lang][0][1]
    v["modesList"] = "".join(f"<li>{x}</li>" for x in S["modesList"][lang])
    v["setList"] = "".join(f"<li>{x}</li>" for x in S["setList"][lang])
    v["langList"] = "".join(f"<span>{x}</span>" for x in S["langList"][lang])
    v["faqList"] = "\n      ".join(f"<details><summary>{q}</summary><p>{a}</p></details>" for q, a in S["faq"][lang])
    v["jsT"] = json.dumps({k: unescape_js(S[k][lang]) for k in JS_KEYS}, ensure_ascii=False)

    html = re.sub(r"\{\{>(\w+)\}\}", lambda m: partials[m.group(1)], tpl)
    missing = set()

    def sub(m):
        k = m.group(1)
        if k not in v:
            missing.add(k)
            return m.group(0)
        return str(v[k])

    html = re.sub(r"\{\{(\w+)\}\}", sub, html)
    if missing:
        raise SystemExit(f"[{lang}] eksik anahtarlar: {sorted(missing)}")
    return html


ROOT_INDEX = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Wortschatz Karten</title>
<link rel="alternate" hreflang="de" href="de/">
<link rel="alternate" hreflang="en" href="en/">
<link rel="alternate" hreflang="tr" href="tr/">
<script>
(function(){var l;try{l=localStorage.getItem('wk-lang')}catch(e){}
var langs=['de','en','tr'];if(langs.indexOf(l)<0){l=(navigator.languages||[navigator.language||'']).map(function(x){return String(x).slice(0,2).toLowerCase()}).filter(function(x){return langs.indexOf(x)>=0})[0]||'de'}
location.replace(l+'/'+location.hash)})();
</script>
<style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#ffd23f;color:#16120f;font:600 18px Arial,sans-serif}nav{display:flex;gap:12px}a{color:inherit;border:3px solid #16120f;border-radius:999px;padding:8px 16px;background:#fff;text-decoration:none}</style>
</head>
<body><nav aria-label="Sprache / Language / Dil"><a href="de/" hreflang="de">Deutsch</a><a href="en/" hreflang="en">English</a><a href="tr/" hreflang="tr">Türkçe</a></nav></body>
</html>
"""

if __name__ == "__main__":
    tpl = (ROOT / "src/template.html").read_text(encoding="utf-8")
    partials = {"defs": (ROOT / "src/_defs.html").read_text(encoding="utf-8"),
                "mascot": (ROOT / "src/_mascot.html").read_text(encoding="utf-8")}
    for lang in LANGS:
        out = OUT / lang / "index.html"
        out.parent.mkdir(exist_ok=True)
        out.write_text(page(lang, tpl, partials), encoding="utf-8")
        print("built:", out.relative_to(ROOT.parent))
    (OUT / "index.html").write_text(ROOT_INDEX, encoding="utf-8")
    print("built: docs/index.html (language redirect)")
