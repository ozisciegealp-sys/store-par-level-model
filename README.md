# Kapasite Planlama ve Otomatik Par Level Modeli

**Türkçe** | [English](README.en.md)

Mağazaların ürün satış raporundan her stok kalemi için **ne kadar sipariş verilmesi gerektiğini**, bu stoğun **depo ve donuk dolap kapasitesine sığıp sığmadığını** ve yeni mağaza planlaması için **minimum raf/dolap ihtiyacını** hesaplayan Excel modeli.

> **Gizlilik ve Sentetik Veri Notu:** Bu çalışma kitabı, 16 mağazalık bölgede kurup 6 bölgeye yaygınlaştırdığım modelin **tamamen sentetik (kurgu) veriyle** yeniden oluşturulmuş gösterim sürümüdür. Kurumsal gizlilik gereği hiçbir marka, gerçek mağaza adı veya şirket verisi yer almaz; ürün reçeteleri, satış rakamları, raf ölçüleri ve sevkiyat takvimi örnekleme amacıyla üretilmiştir.

## Hangi problemi çözüyor

Bir mağazanın sipariş miktarı üç değişkene bağlıdır: ürünlerin satış hızı, her ürünün hangi sarf/stok kalemini ne kadar tükettiği ve bir sonraki sevkiyata kaç gün olduğu. Model bu üç halkayı tek zincirde birleştirir:

- Satış **ürün** bazında gelir (latte, kruvasan…), sipariş ise **stok kalemi** bazında verilir (bardak, kapak, kahve çekirdeği, donuk ürün…). Reçete matrisi bu dönüşümü otomatik yapar.
- Sipariş miktarı, bir sonraki sevkiyata kadar geçen gün sayısına ve teslimat vardiyasına (`SABAH` / `AKŞAM`) göre ölçeklenir.
- Hesaplanan par stoğunun koli hacmi (`m³`), mağazanın kuru depo raf ve donuk dolap kapasitesiyle karşılaştırılır; ayrıca iş hacmine göre gereken minimum ünite adedi çıkarılarak yeni mağaza planlamasına girdi sağlanır.

## Öncesi

Her mağazanın siparişi stok kalemi bazında tek tek elle hesaplanıyordu. Model önce 16 mağazalık bölgede devreye alındı, ardından bölge yönetimi aracılığıyla 6 bölgeye yaygınlaştırıldı.

## Nasıl çalışıyor

```
Satis_Raporu ──────► Receteler ──────► Par_Level ──────► Kapasite & Kapasite_Ozeti
 (son 30 gün /        (ürün satışı      (tüketim ×        (koli hacmi ÷ raf ve donuk
  son 7 gün satış)     × reçete =        periyot × %115    kapasitesi + gerekli min.
                       günlük tüketim)   − sayım = koli)   ünite sayısı)
       │
       └───────────► Bolge_Ozet (5 farklı mağaza formatının satış karması)
```

| Sayfa | Ne yapar |
|---|---|
| `Nasıl Kullanılır` | Akış, renk kodu ve periyot günü hesabının açıklaması |
| `Ayarlar` | Mağaza, sipariş/sevkiyat tarihleri ve rapor dönemi (`Aylık (30 gün)` / `Haftalık (7 gün)`) seçimi |
| `Par_Level` | Stok kalemi bazında günlük tüketim, par seviyesi, eldeki sayım ve koli bazında sipariş |
| `Kapasite` | Seçili mağazanın kuru depo ve donuk dolap doluluk oranı ile önerilen minimum ünite sayısı |
| `Kapasite_Ozeti` | 5 mağaza formatının (AVM, Cadde, Kampüs, Havalimanı, Otoyol) yan yana kapasite ve ünite karşılaştırması |
| `Magazalar` | Mağaza formatı, rota, teslimat saati/vardiyası, mevcut raf ve donuk kapı sayıları |
| `Stok_Kalemleri` | Koli içi miktar, koli hacmi (`m³`) ve depolama alanı (`Depo` / `Donuk`) |
| `Sevkiyat_Takvimi` | Mağaza bazında haftalık depo ve donuk sevkiyat günleri |
| `Satis_Raporu` | POS/ERP'den alınan son 30 günlük ve son 7 günlük ürün satış verisi |
| `Receteler` | Ürün × stok kalemi kullanım matrisi ve günlük tüketim toplamları |
| `Bolge_Ozet` | Tüm mağazaların ürün grubu bazında satış dağılımı |

### Temel formüller

- **Günlük satış** (`Receteler`, C sütunu): `Ayarlar!B9` seçimine göre **son 30 günlük** (varsayılan, genel sipariş ve kapasite planlama için) veya **son 7 günlük** (bayram/kampanya gibi kısa talep değişimleri için) satış toplamı gün sayısına bölünür.
- **Günlük tüketim** (`Receteler`, toplam satırı): `SUMPRODUCT(günlük satış; birim reçete kullanımı)`
- **Periyot günü** (`Ayarlar`): `sonraki sevkiyat − sipariş tarihi + vardiya düzeltmesi` (öğleden sonra `AKŞAM` teslimatı alan mağazalara o günün satışı sevkiyattan önce gerçekleştiği için `+1` gün eklenir).
- **Par seviyesi**: `günlük tüketim × periyot × (1 + güvenlik stoğu)` (varsayılan güvenlik stoğu `%15`)
- **Sipariş (koli)**: `YUKARIYUVARLA(MAX(0; par − sayım) ÷ koli içi miktar; 0)`
- **Doluluk ve ünite hesabı**: `par seviyesindeki koli hacmi ÷ (ünite adedi × ünite hacmi)`

## Kurulum ve kullanım

Makro içermez; Excel 2010+ ve LibreOffice'te doğrudan açılır.

1. `par-level-modeli.xlsx` dosyasını açın.
2. `Ayarlar` sayfasındaki sarı hücrelerden mağazayı, sipariş/sevkiyat tarihlerini ve rapor dönemini (`Aylık (30 gün)` veya `Haftalık (7 gün)`) seçin.
3. `Par_Level` sayfasında **Sayım (eldeki)** sütununu güncelleyip **Sipariş (koli)** çıktısını alın.
4. `Kapasite` ve `Kapasite_Ozeti` sayfalarından depo/donuk doluluk oranlarını ve ünite yeterliliğini kontrol edin.

Renk kodu: **mavi** yazı = sabit girdi, **sarı** zemin = kullanımda seçilen/girilen hücre, **yeşil** yazı = sayfa içi referans, **siyah** yazı = formül.

## Lisans

MIT. Ayrıntılar için `LICENSE` dosyasına bakın.
