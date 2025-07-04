# YEV2MF - XML'den Excel'e Fiş Dönüştürme Aracı

Bu araç, XML formatındaki yevmiye defyerinden muhasebe fişlerini Excel formatına dönüştürmek için tasarlanmıştır. Özellikle büyük veri setlerini işlemek için optimize edilmiştir.
Temel amaç yevmiye defterindeki verileri luca programına yüklemektir.
Oluşan excel dosyalarını luca içerisinde Muhasebe -> Fiş İşlemleri -> Excel Veri Aktarımı kısmından aktarabilirsiniz.

## Özellikler

- XML formatındaki fişleri Excel'e dönüştürme
- Büyük veri setlerini 50'şer fişlik gruplara ayırma
- Otomatik tarih formatı dönüşümü
- Hata yönetimi ve loglama

  
## Gereksinimler

- Python 3.6 veya üzeri
- Aşağıdaki Python kütüphaneleri:
  - openpyxl
  - xml.etree.ElementTree (standart kütüphane)
  
## Kurulum/Kullanım

```bash
   git clone [https://github.com/kullanici_adiniz/YEV2MF.git](https://github.com/kullanici_adiniz/YEV2MF.git)
   cd YEV2MF
```

1. Dönüştürmek istediğiniz XML dosyalarını XMLYevmiye klasörüne koyun.

2. Aşağıdaki komutu çalıştırın:

```bash
python yev2mf.py
```
3. Dönüştürülen dosyalar ExcelMuhasebeFisi klasörüne kaydedilecektir.

## Kurulum/Kullanım 2

1. CalistirYEV2MF.bat dosyasını yönetici olarak çalıştırın.

## Ekran Görüntüleri
![Uygulama Ekran Görüntüsü](https://yavuzselimkilic.com/araclar/YEV2MF/EkranGoruntusu/Program.png)

![Uygulama Ekran Görüntüsü](https://yavuzselimkilic.com/araclar/YEV2MF/EkranGoruntusu/Excel.png)

![Uygulama Ekran Görüntüsü](https://yavuzselimkilic.com/araclar/YEV2MF/EkranGoruntusu/Luca.png)
