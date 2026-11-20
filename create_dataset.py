import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Dengeli 100 Soruluk Veri Seti (Kolay: 1-35, Orta: 36-70, Zor: 71-100)
questions_data = [
    # =========================================================
    # KOLAY SORULAR (1-35): Temel Genel Kültür, Fen, Coğrafya, Basit Matematik
    # =========================================================
    ("Coğrafya", "Türkiye'nin başkenti neresidir?", "Ankara"),
    ("Fen Bilimleri", "Su deniz seviyesinde ve normal atmosfer basıncında kaç santigrat derecede kaynar?", "100"),
    ("Astronomi", "Güneş sistemindeki en büyük gezegen hangisidir?", "Jüpiter"),
    ("Edebiyat / Tarih", "İstiklal Marşı'nın şairi kimdir?", "Mehmet Akif Ersoy"),
    ("Genel Kültür", "Bir yılda kaç gün vardır (artık yıllar hariç)?", "365"),
    ("Astronomi", "Dünya'nın tek doğal uydusunun adı nedir?", "Ay"),
    ("Coğrafya", "Türkiye'nin en yüksek dağı hangisidir?", "Ağrı Dağı"),
    ("Kimya", "Kimyada suyun formülü nedir?", "H2O"),
    ("Coğrafya", "Fransa'nın başkenti neresidir?", "Paris"),
    ("Matematik", "Bir üçgenin iç açılarının toplamı kaç derecedir?", "180"),
    ("Ekonomi", "Türk Lirası'nın uluslararası para birimi kodu (kısaltması) nedir?", "TRY"),
    ("Biyoloji", "İnsan vücudunda kanı pompalayan organ hangisidir?", "Kalp"),
    ("Fen Bilimleri", "Gökkuşağında kaç ana renk vardır?", "7"),
    ("Coğrafya", "Türkiye'nin kaç coğrafi bölgesi vardır?", "7"),
    ("Bilim Tarihi", "Telefonu icat eden bilim insanı kimdir?", "Alexander Graham Bell"),
    ("Coğrafya", "Japonya'nın başkenti neresidir?", "Tokyo"),
    ("Matematik", "Bir saat kaç dakikadır?", "60"),
    ("Coğrafya", "Dünya üzerindeki en büyük okyanus hangisidir?", "Büyük Okyanus (Pasifik Okyanusu)"),
    ("Tarih", "Mustafa Kemal Atatürk hangi yılda doğmuştur?", "1881"),
    ("Spor", "Futbol maçlarında bir takım sahaya kaç oyuncu ile çıkar?", "11"),
    ("Biyoloji", "Bitkilerin güneş ışığı kullanarak besin üretmesi olayına ne ad verilir?", "Fotosentez"),
    ("Coğrafya", "İngiltere'nin başkenti neresidir?", "Londra"),
    ("Matematik", "8 kere 7 kaç eder?", "56"),
    ("Tarih", "Türkiye Cumhuriyeti hangi tarihte (gün, ay, yıl) ilan edilmiştir?", "29 Ekim 1923"),
    ("Sanat", "Mona Lisa tablosunun ressamı kimdir?", "Leonardo da Vinci"),
    ("Kimya", "Periyodik tabloda 'O' sembolü hangi elementi temsil eder?", "Oksijen"),
    ("Matematik", "En küçük asal sayı kaçtır?", "2"),
    ("Coğrafya", "İtalya'nın başkenti neresidir?", "Roma"),
    ("Matematik", "Bir haftada toplam kaç saat vardır?", "168"),
    ("Coğrafya", "Dünya'nın kutuplardan basık, ekvatordan şişkin olan özel şekline ne ad verilir?", "Geoit"),
    ("Coğrafya", "Türkiye'nin yüzölçümü bakımından en büyük ili hangisidir?", "Konya"),
    ("Fizik", "Yerçekimi kanununu keşfeden ünlü İngiliz bilim insanı kimdir?", "Isaac Newton"),
    ("Matematik", "Bir düzinede kaç adet nesne bulunur?", "12"),
    ("Biyoloji", "Kuzey Kutbu'nda yaşayan, beyaz kürküyle bilinen büyük ayı türü nedir?", "Kutup ayısı"),
    ("Matematik", "250 sayısının %20'si kaçtır?", "50"),

    # =========================================================
    # ORTA SORULAR (36-70): Kültür, Fen, Tarih ve Orta Düzey Matematik Problemleri
    # =========================================================
    ("Kimya", "Periyodik tabloda 'Au' sembolü ile gösterilen element hangisidir?", "Altın"),
    ("Edebiyat", "Suç ve Ceza romanının yazarı kimdir?", "Fyodor Dostoyevski"),
    ("Tarih", "Osmanlı Devleti hangi yılda kurulmuştur?", "1299"),
    ("Biyoloji", "Hücrenin enerji santrali olarak bilinen organel hangisidir?", "Mitokondri"),
    ("Astronomi", "Güneş'e en yakın gezegen hangisidir?", "Merkür"),
    ("Tarih", "İstanbul'un Fethi hangi yılda gerçekleşmiştir?", "1453"),
    ("Edebiyat", "İnce Memed romanının yazarı kimdir?", "Yaşar Kemal"),
    ("Fizik / Kimya", "Atomun çekirdeğinde bulunan yüksüz (nötr) parçacığa ne ad verilir?", "Nötron"),
    ("Biyoloji", "İnsan vücudundaki en uzun ve en güçlü kemik hangisidir?", "Uyluk kemiği (Femur)"),
    ("Coğrafya", "Avustralya'nın başkenti neresidir?", "Kanberra (Canberra)"),
    ("Matematik", "3x + 7 = 28 denkleminde x kaçtır?", "7"),
    ("Edebiyat", "Kürk Mantolu Madonna adlı eserin yazarı kimdir?", "Sabahattin Ali"),
    ("Coğrafya / Kimya", "Dünya atmosferinde en yüksek oranda (%78) bulunan gaz hangisidir?", "Azot (Nitrojen)"),
    ("Tarih", "İlk Türk kadın savaş pilotu kimdir?", "Sabiha Gökçen"),
    ("Bilgisayar Bilimi", "Bilgisayar bilimlerinde 1 Kilobyte (ikili sistemde / GiB-KiB standardında) kaç Byte'tır?", "1024"),
    ("Coğrafya", "Brezilya'nın resmi dili nedir?", "Portekizce"),
    ("Edebiyat", "Sefiller (Les Misérables) romanının yazarı kimdir?", "Victor Hugo"),
    ("Tarih", "Türkiye'de kadınlara milletvekili seçme ve seçilme hakkı hangi yılda verilmiştir?", "1934"),
    ("Matematik", "Bir dik üçgende dik kenarların uzunlukları 6 cm ve 8 cm ise hipotenüs uzunluğu kaç cm'dir?", "10"),
    ("Tarih", "Fatih Sultan Mehmet'in babası olan Osmanlı padişahı kimdir?", "II. Murad"),
    ("Bilim Tarihi", "Penisilini keşfeden İskoç bilim insanı kimdir?", "Alexander Fleming"),
    ("Matematik", "Saatte 90 km sabit hızla giden bir araç 40 dakikada kaç kilometre yol alır?", "60"),
    ("Fizik", "Elektrik akımının şiddetini ölçmek için kullanılan birimin adı nedir?", "Amper"),
    ("Coğrafya", "Türkiye sınırları içindeki uzunluğuna göre Türkiye'nin en uzun nehri hangisidir?", "Kızılırmak"),
    ("Edebiyat", "Saatleri Ayarlama Enstitüsü romanının yazarı kimdir?", "Ahmet Hamdi Tanpınar"),
    ("Matematik", "Bir ürünün fiyatına önce %20 zam, ardından zamlı fiyat üzerinden %20 indirim yapılırsa başlangıç fiyatına göre yüzde kaç kar veya zarar edilir?", "%4 zarar (veya %4 indirim)"),
    ("Tarih", "Magna Carta (Büyük Sözleşme) hangi yılda imzalanmıştır?", "1215"),
    ("Astronomi", "Güneş sisteminde kendi ekseni etrafında diğer gezegenlerin tersi yönünde dönen ve en sıcak gezegen olan hangisidir?", "Venüs"),
    ("Bilim Tarihi", "Kuduz aşısını bulan Fransız bilim insanı kimdir?", "Louis Pasteur"),
    ("Matematik", "Ardışık 5 tek tam sayının toplamı 135 olduğuna göre bu sayıların en büyüğü kaçtır?", "31"),
    ("Biyoloji", "DNA'nın çift sarmal yapısını keşfeden bilim insanları James Watson ve kimdir?", "Francis Crick"),
    ("Matematik", "Logaritma 2 tabanında 64 (log₂64) işleminin sonucu kaçtır?", "6"),
    ("Tarih", "Lozan Barış Antlaşması hangi tarihte (gün, ay, yıl) imzalanmıştır?", "24 Temmuz 1923"),
    ("Kimya", "Oda sıcaklığında (25 °C) sıvı halde bulunan tek metal element hangisidir?", "Cıva"),
    ("Matematik", "Bir sınıftaki öğrencilerin %40'ı kızdır. Sınıfta 18 erkek öğrenci olduğuna göre toplam öğrenci sayısı kaçtır?", "30"),

    # =========================================================
    # ZOR SORULAR (71-100): Çok Adımlı Matematik, Mantık, Bilim ve Detaylı Kültür
    # =========================================================
    ("Matematik", "(17 × 24) - (13 × 18) işleminin sonucu kaçtır?", "174"),
    ("Matematik / Mantık", "Bir kalem ve bir silginin toplam fiyatı 110 TL'dir. Kalem, silgiden 100 TL daha pahalı olduğuna göre silginin fiyatı kaç TL'dir?", "5 TL"),
    ("Mantık / Matematik", "5 makine 5 dakikada 5 düğme üretiyorsa, 100 makine 100 düğmeyi kaç dakikada üretir?", "5 dakika"),
    ("Matematik / Mantık", "Bir göldeki nilüfer çiçekleri her gün iki katına çıkmaktadır. Nilüferlerin tüm gölü kaplaması 48 gün sürdüğüne göre, gölün dörtte birini (1/4'ünü) kaplaması kaçıncı günde gerçekleşir?", "46. gün"),
    ("Matematik", "Bir işi Ahmet tek başına 12 günde, Mehmet ise tek başına 24 günde bitirebilmektedir. İkisi birlikte aynı işi kaç günde bitirirler?", "8"),
    ("Matematik / Geometri", "Saat tam 04:20'yi gösterirken akrep ile yelkovan arasındaki küçük açı kaç derecedir?", "10 derece"),
    ("Matematik / Kombinasyon", "7 kişilik bir gruptan 3 kişilik bir komisyon kaç farklı şekilde seçilebilir?", "35"),
    ("Matematik / Olasılık", "İki zar aynı anda havaya atıldığında üst yüze gelen sayıların toplamının 9 olma olasılığı en sade kesir olarak kaçtır?", "1/9"),
    ("Matematik / Mantık", "Bir salyangoz 14 metre derinliğindeki bir kuyunun dibindedir. Gündüzleri 4 metre yukarı tırmanmakta, geceleri ise 2 metre aşağı kaymaktadır. Salyangoz kuyunun tepesine kaçıncı günde ulaşır?", "6. gün"),
    ("Tarih", "Osmanlı Devleti'nde ilk resmi gazete olan Takvim-i Vekayi hangi padişah döneminde ve hangi yılda çıkarılmıştır?", "II. Mahmud - 1831"),
    ("Matematik", "2³ + 3⁴ - 4² işleminin sonucu kaçtır?", "73"),
    ("Fizik", "Kütlesi 4 kg olan bir cismin hızı 5 m/s olduğuna göre kinetik enerjisi kaç Joule'dür?", "50 Joule"),
    ("Coğrafya", "Türkiye'de dört coğrafi bölgede (İç Anadolu, Karadeniz, Marmara ve Ege) toprakları bulunan tek ilimiz hangisidir?", "Bilecik"),
    ("Matematik", "Bir baba 38, oğlu ise 10 yaşındadır. Kaç yıl sonra babanın yaşı oğlunun yaşının tam 3 katı olur?", "4 yıl sonra"),
    ("Coğrafya", "Türkiye'nin komşuları arasında kara sınırımızın en kısa olduğu ülke hangisidir?", "Azerbaycan (Nahçıvan)"),
    ("Matematik", "1'den 20'ye kadar (20 dahil) olan tüm asal sayıların toplamı kaçtır?", "77"),
    ("Edebiyat", "Türk edebiyatında ilk psikolojik roman kabul edilen 'Eylül' adlı eserin yazarı kimdir?", "Mehmet Rauf"),
    ("Biyoloji", "İnsan vücudunda hem hormon (endokrin) hem de sindirim enzimi (ekzokrin) salgılayan karma bez organ hangisidir?", "Pankreas"),
    ("Matematik / Olasılık", "İki çocuklu bir ailede çocuklardan en az birinin kız olduğu biliniyorsa, her iki çocuğun da kız olma olasılığı kaçtır?", "1/3"),
    ("Tarih", "Osmanlı İmparatorluğu'nun ilk anayasası olan ve 1876 yılında ilan edilen anayasanın adı nedir?", "Kanun-i Esasi"),
    ("Kimya", "Ametaller arasında oda koşullarında elektriği iyi ileten ve kalem uçlarında da kullanılan karbon allotropu hangisidir?", "Grafit"),
    ("Matematik", "Bir sayının 3 katının 14 eksiği, aynı sayının yarısının 21 fazlasına eşittir. Bu sayı kaçtır?", "14"),
    ("Fizik", "Suyun yoğunluğunun (özkütlesinin) en yüksek olduğu sıcaklık kaç santigrat derecedir?", "+4 °C"),
    ("Matematik", "Bir araç gideceği yolun ilk yarısını saatte 60 km hızla, ikinci yarısını ise saatte 120 km hızla gitmiştir. Bu aracın tüm yol boyunca ortalama hızı saatte kaç km'dir?", "80 km/s"),
    ("Bilim Tarihi", "2015 yılında 'DNA onarımı' hakkındaki çalışmalarıyla Nobel Kimya Ödülü'nü kazanan Türk bilim insanı kimdir?", "Aziz Sancar"),
    ("Matematik", "Taban yarıçapı 3 cm, yüksekliği 4 cm olan dik dairesel koninin yanal ayrıtı (ana doğrusu) kaç cm'dir?", "5"),
    ("Edebiyat / Tarih", "Divanü Lügati't-Türk adlı ilk Türkçe sözlük ve dil bilgisi kitabının yazarı kimdir?", "Kaşgarlı Mahmud"),
    ("Matematik", "1000 TL anapara, yıllık %20 bileşik faiz oranıyla 2 yıl sonunda toplam kaç TL olur?", "1440 TL"),
    ("Felsefe", "Felsefede 'Düşünüyorum, öyleyse varım' (Cogito, ergo sum) sözü hangi filozofa aittir?", "René Descartes"),
    ("Matematik", "Bir torbada 4 kırmızı ve 6 mavi top vardır. Torbadan geri konulmamak şartıyla rastgele arka arkaya çekilen 2 topun da kırmızı olma olasılığı en sade kesir olarak kaçtır?", "2/15")
]

rows = []
for idx, (cat, q, a) in enumerate(questions_data, start=1):
    rows.append({
        "Soru No": idx,
        "Kategori": cat,
        "Soru": q,
        "Ana Cevap": a
    })

df = pd.DataFrame(rows)

def format_excel(filepath):
    wb = load_workbook(filepath)
    ws = wb.active
    ws.title = "Soru Dataset"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    data_font = Font(name="Calibri", size=11, color="000000")
    bold_font = Font(name="Calibri", size=11, bold=True, color="1F4E78")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws.row_dimensions[1].height = 28
    for col_idx in range(1, 5):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row_idx in range(2, len(df) + 2):
        ws.row_dimensions[row_idx].height = 24
        row_fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, 5):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.border = thin_border
            if col_idx == 1:
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 2:
                cell.font = data_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 3:
                cell.font = data_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_idx == 4:
                cell.font = Font(name="Calibri", size=11, bold=True, color="000000")
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 85
    ws.column_dimensions["D"].width = 35

    ws.freeze_panes = "A2"
    wb.save(filepath)

target_file = "soru_dataset.xlsx"
try:
    df.to_excel(target_file, index=False)
    format_excel(target_file)
    print(f"Başarılı: '{target_file}' güncellendi ve biçimlendirildi ({len(df)} soru).")
except PermissionError:
    alt_file = "soru_dataset_yeni.xlsx"
    df.to_excel(alt_file, index=False)
    format_excel(alt_file)
    print(f"Not: '{target_file}' açık olduğu için '{alt_file}' olarak kaydedildi.")
