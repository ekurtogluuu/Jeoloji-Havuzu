# Jeoloji Havuzu

Istanbul jeoloji raporundan zemin etudu programlarinda kullanilabilecek
yapilandirilmis jeolojik birim katalogu uretmek icin hazirlanan veri cikarma
projesi.

## Amac

Bu proje, Istanbul Buyuksehir Belediyesi tarafindan yayimlanan Istanbul jeoloji
raporundaki formasyon, uye, grup ve benzeri jeolojik birimleri DOCX kaynaktan
okuyarak JSON ve CSV formatinda denetlenebilir bir veri setine donusturur.

Uretilen katalog, zemin etudu yazilimlarinda jeoloji bolumu icin arama, secim,
on doldurma ve kontrol veri kaynagi olarak kullanilmak uzere tasarlanmistir.

## Proje Yapisi

```text
Jeoloji-Havuzu/
|-- src/
|   `-- extract_units.py
|-- output/
|   |-- geology_units.json
|   |-- geology_units.csv
|   `-- extraction_report.md
|-- README.md
|-- .gitignore
`-- .gitattributes
```

## Kullanim

Gerekli Python paketleri:

```bash
pip install python-docx
```

Veri setini yeniden uretmek:

```bash
python src/extract_units.py
```

Soz dizimi kontrolu:

```bash
python -m py_compile src/extract_units.py
```

## Ciktilar

- `output/geology_units.json`: Program entegrasyonu icin yapilandirilmis ana veri.
- `output/geology_units.csv`: Excel veya tablo araclariyla kontrol edilebilir denetim ciktisi.
- `output/extraction_report.md`: Kaynak, toplam kayit, alan doluluk oranlari ve manuel kontrol listesi.

Mevcut veri seti 43 jeolojik birim icerir. Ciktilarda ham baslik, temizlenmis
baslik, DOCX paragraf araligi, litoloji, dagilim, dokanak, kalinlik, fosil/yas,
korelasyon ve kalite alanlari bulunur.

## Kaynak Belgeler

Kaynak PDF ve DOCX dosyalari buyuk boyutlu oldugu ve kullanim/telif durumlari
ayrica degerlendirilmesi gerektigi icin repoda tutulmamasi onerilir. Betik
calismak icin proje kokunde bir `.docx` kaynak dosyasi bekler.

## Veri Kalitesi

Bu proje otomatik veri cikarma hattidir. Uretilen katalog muhendislik
kullanimindan once manuel orneklem kontroluyle dogrulanmalidir. Kalite takibi
icin `extraction_report.md` dosyasindaki alan doluluk oranlari, eksik bolumler
ve manuel kontrol listesi esas alinmalidir.

## Lisans Notu

Kod icin acik kaynak lisans eklenebilir. Kaynak rapor ve rapordan turetilen veri
icin kullanim haklari ayrica degerlendirilmelidir.
