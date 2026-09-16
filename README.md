# Par Level ve Kapasite Planlama Modeli

Mağazanın ürün satış raporundan, her stok kalemi için **ne kadar sipariş verilmesi gerektiğini** ve bu stoğun **depoya sığıp sığmadığını** hesaplayan Excel modeli.

> **Not:** Bu dosya, bir kahve zincirinin bölge operasyonunda kurduğum ve kullandığım modelin, **sahte veriyle yeniden oluşturulmuş gösterim sürümüdür**. Orijinal dosya şirket verisi içerdiği için paylaşılmamıştır. Mağaza adları, ürünler, satış adetleri, raf ölçüleri ve sevkiyat takvimi uydurmadır. Formül mimarisi orijinaliyle aynıdır.

## Hangi problemi çözüyor

Bir mağazanın sipariş vereceği miktar üç şeye bağlıdır: ürünlerin ne kadar sattığı, her ürünün hangi sarf malzemesini ne kadar tükettiği ve bir sonraki sevkiyata kaç gün olduğu. Bu model üçünü tek zincirde birleştirir:

- Satış **ürün** bazında gelir (latte, kruvasan…), sipariş ise **stok kalemi** bazında verilir (büyük kap, kapak, kahve çekirdeği…). Reçete matrisi bu dönüşümü yapar.
- Sipariş miktarı, bir sonraki sevkiyata kadar geçen gün sayısına (periyot) göre ölçeklenir.
- Hesaplanan stoğun kaç koli ettiği ve bu kolilerin mağazanın raf ve donuk dolap hacmine oranı çıkarılır. Böylece stok seviyesi fiziksel kapasiteyle birlikte planlanır.

## Öncesi

Her mağazanın siparişi, stok kalemi bazında tek tek elle hesaplanıp veriliyordu. Model kurulduktan sonra bölgedeki 16 mağazada kullanıldı ve bazı mağazalar siparişlerini bu modelin çıktısıyla vermeye geçti. Model daha sonra bölge müdürü aracılığıyla altı bölgeye yaygınlaştırıldı.

## Nasıl çalışıyor

```
Satis_Raporu ──► Receteler ──► Par_Level ──► Kapasite
 (haftalık       (ürün satışı    (tüketim ×       (koli hacmi ÷
  ürün satışı)    × reçete =      periyot ×        raf/dolap
                  günlük          %115 − sayım     kapasitesi)
                  tüketim)        = sipariş)
       │
       └──────► Bolge_Ozet (mağazaların satış karması)
```

| Sayfa | Ne yapar |
|---|---|
| `Nasıl Kullanılır` | Akış, renk kodu ve periyot hesabının açıklaması |
| `Ayarlar` | Mağaza, sipariş tarihi ve sonraki sevkiyat tarihi seçilir; periyot günü burada hesaplanır |
| `Par_Level` | Stok kalemi bazında par seviyesi, sipariş miktarı ve koli sayısı |
| `Kapasite` | Depo ve donuk dolap doluluk oranı |
| `Magazalar` | Mağaza bilgileri: format, rota, teslimat saati, raf ve dolap adedi |
| `Stok_Kalemleri` | Koli içi miktar, koli hacmi, depolama yeri |
| `Sevkiyat_Takvimi` | Hangi mağazaya hangi gün hangi sevkiyatın geldiği |
| `Satis_Raporu` | POS/ERP'den alınan haftalık satış raporunun yapıştırıldığı sayfa |
| `Receteler` | Ürün × stok kalemi kullanım matrisi ve günlük tüketim satırı |
| `Bolge_Ozet` | Tüm mağazaların ana grup bazında satış dağılımı |

### Temel formüller

- **Günlük satış** (`Receteler`, C sütunu): seçili mağazanın rapordaki satışı `SUMIFS` ile çekilir, raporun kapsadığı gün sayısına bölünür.
  - Genel sipariş planı ve yeni mağaza kapasite hesabı için **aylık** satış raporu kullanılır (`Ayarlar` B10 = 30).
  - Bayram gibi talebin kısa sürede değiştiği dönemlerde **haftalık** rapor kullanılır (`Ayarlar` B10 = 7). Örnek dosya bu ayarla gelir.
- **Günlük tüketim** (`Receteler`, toplam satırı): `SUMPRODUCT(günlük satış; kullanım miktarı)`
- **Periyot günü** (`Ayarlar`): `sonraki sevkiyat − sipariş tarihi + vardiya düzeltmesi`. Sevkiyatı öğleden sonra alan mağazaya +1 gün eklenir, çünkü o günün satışı sevkiyat gelmeden yapılır.
- **Par seviyesi**: `günlük tüketim × periyot × (1 + güvenlik stoğu)`, güvenlik stoğu %15
- **Sipariş**: `MAX(0; par − sayım)`, koli içi miktara bölünüp yukarı yuvarlanır
- **Doluluk**: `par seviyesindeki koli hacmi ÷ (ünite adedi × ünite hacmi)`

## Kurulum ve kullanım

Makro yoktur, kurulum gerekmez. Dosya Excel 2010 ve sonrası ile LibreOffice'te açılır.

1. `par-level-modeli.xlsx` dosyasını indirip açın.
2. `Ayarlar` sayfasında sarı hücreleri doldurun: mağaza (listeden), sipariş tarihi, bir sonraki sevkiyat tarihi. Sonraki sevkiyat tarihi için `Sevkiyat_Takvimi` sayfasına bakın.
3. `Par_Level` sayfasında **Sayım** sütununa eldeki miktarları girin.
4. **Sipariş (koli)** sütunu verilecek siparişi gösterir. `Kapasite` sayfasında stoğun alana sığıp sığmadığını kontrol edin.

Yeni dönemin satışını kullanmak için raporu `Satis_Raporu` sayfasına aynı sütun düzeniyle yapıştırmanız yeterlidir.

## Kendi verinize uyarlama

| Nerede | Ne girilir |
|---|---|
| `Magazalar` | Mağaza adı, format, teslimat vardiyası (`SABAH` / `AKŞAM`), raf ve donuk dolap kapı adedi |
| `Stok_Kalemleri` | Kalem adı, koli içi miktar, koli hacmi (m³), depolama yeri (`Depo` / `Donuk`) |
| `Receteler` | Satır: ürün. Sütun: stok kalemi. Hücre: bir satışta kullanılan miktar |
| `Sevkiyat_Takvimi` | Mağaza bazında haftalık sevkiyat günleri |
| `Ayarlar` B10 / B11 | Satış raporunun kaç günü kapsadığı (haftalık rapor 7, aylık rapor 30), güvenlik stoğu oranı |
| `Kapasite` B4 / B5 | Bir raf ünitesinin ve bir donuk dolap kapısının hacmi |

Yeni stok kalemi eklerken aynı adı hem `Stok_Kalemleri`'ne hem `Receteler`'in başlık satırına yazın. `Par_Level` kalemi ada göre eşleştirir. Satır veya sütun eklendiğinde `Par_Level` ve `Receteler` toplam satırındaki formül aralıklarını genişletin.

Renk kodu: **mavi** yazı elle girilen değer, **sarı** zemin her kullanımda doldurulan hücre, **yeşil** yazı başka sayfadan gelen değer, **siyah** yazı formüldür.

## Sınırlar

- Sonraki sevkiyat tarihi elle seçilir; takvimden otomatik bulunmaz.
- Talep, seçilen rapor döneminin günlük ortalamasıdır. Özel dönem etkisi ancak o döneme ait haftalık rapor yüklenerek yakalanır; model kendi başına tahmin yapmaz.
- Raf ve dolap hacimleri örnek değerdir; gerçek kullanımda mağazada ölçülmelidir.

## Lisans

MIT. Ayrıntılar için `LICENSE` dosyasına bakın.
