# Mağaza Gün Sonu (Mark-Out), Çözündürme ve Fire Analiz Sistemi

**Türkçe** | [English](README.en.md)

Çok şubeli perakende ve F&B mağazalarında donuk/günlük yiyeceklerin gün sonu sayımını, ERP (Oracle) satış raporuyla mutabakatını, adisyon (ADT/UPH) bazlı çözündürme planlamasını, SKT takibini ve aylık fire analizini yöneten Excel + VBA çalışma kitabı.

> **Gizlilik ve Sentetik Veri Notu:** Bu depo, bölge operasyonundaki 16 mağazada kullanılan gün sonu raporlama ve stok kontrol sisteminin mağaza ayağını **tamamen sentetik (kurgu) veriyle** gösterir. Kurumsal gizlilik gereği marka, mağaza ve şirket verileri yer almaz; ürün adları, SKU kodları, maliyetler ve satış adetleri örnekleme amacıyla üretilmiştir.
>
> Orijinal akışta günlük satış raporu ERP sisteminden (Oracle) indirilip makroyla tek tuşta bu forma okutuluyordu. Bu depoda ERP çıktısını `ornek-satis-raporu.xlsx` temsil eder.

## Hangi problemi çözüyor

Her akşam kapanışta ve ertesi gün hazırlığında şu soruların hızlı ve hatasız yanıtlanması gerekir: Donukta ve vitrinde ne kadar ürün kaldı? Gün içinde ne kadar satıldı, ne kadar fire (mark-out) oluştu? Sayılan stok beklenenle tutuyor mu? Yarınki tahmini müşteri sayısına göre kaç adet ürün çözündürülmeli ve hangi SKT etiketi basılmalı?

Bu sistemde:

- Satış adetleri ERP raporundan tek kısayolla (`Ctrl+Shift+S`) SKU bazında forma çekilir.
- Donuk ve günlük beklenen kapanış ile sayım farkları (kayıp/fire) otomatik hesaplanır.
- Yarınki tahmini adisyona (ADT) ve ürünlerin 100 adisyon başına satış hızına (UPH) göre çözündürme önerisi ve donuk stok yeterlilik uyarısı üretilir.
- Çözündürmeye alınan ve dünden devreden ürünlerin SKT'leri ayrı sayfada otomatik hesaplanır.
- 30 günlük satış ve fire verisinden **Yıldız Ürünler** (yüksek satış, düşük fire) ile **Sorunlu Ürünler** (hedef üstü fire) sınıflandırılır.
- Gün kapatıldığında (`Ctrl+Shift+K`) kapanış sayımları yarının açılış stoğuna devreder ve günün kaydı arşive işlenir.

## Öncesi

Satış adetleri ERP raporundan ürün ürün bakılarak elle yazılıyor, kapanış farkları ve çözündürme adetleri manuel hesaplanıyor, ertesi günün devir stokları elle taşınıyordu (~60 dk/gün → ~1 dk/gün).

## Dosyalar

| Dosya | İçerik |
|---|---|
| `markout-formu.xlsx` | Ana çalışma kitabı (formüller, analiz sayfaları ve sentetik örnek veri hazır) |
| `markout.bas` | Satış çekme, PDF alma ve gün kapatma/devir işlemlerini yürüten VBA modülü |
| `ornek-satis-raporu.xlsx` | Makroyu denemek için sentetik günlük ERP satış raporu ve dönem özeti |

`markout.bas` içinde bilerek ASCII karakterler kullanılmıştır; VBA düzenleyicisi `.bas` dosyalarını sistemin ANSI kodlamasıyla okuduğu için kod hem GitHub'da hem de Excel'e aktarıldığında bozulmadan çalışır. Formül ve analiz sayfaları makro eklenmeden de doğrudan incelenebilir.

## Sayfalar

| Sayfa | Ne yapar |
|---|---|
| `Nasil_Kullanilir` | Günlük operasyon akışı, hesap mantığı ve renk kodu |
| `Markout` | Ana form: donuk ve günlük stok hareketleri, beklenen kapanış ve sayım farkı |
| `SKT` | Bugün çözündürmeye alınan ve dünden devreden ürünlerin saat ve SKT tablosu |
| `Cozundurme_Plani` | Yarınki tahmini adisyon (ADT) ve UPH hızına göre ürün bazlı çözündürme önerisi ve donuk stok kontrolü |
| `Ayarlar` | Mağaza adı, ortalama baz seçimi, çözündürme/mark-out saatleri, hedef fire oranı ve ERP rapor sütun ayarları |
| `Satis_Ortalama` | Son 7 ve 30 günlük satıştan ortalama günlük satış hesabı |
| `Fire_Analizi` | 30 günlük satış ve atık verisinden fire maliyeti ile Yıldız / Dengeli / Sorunlu ürün sınıflandırması |
| `Bos_Form` | Sahada kapanış sayımı için yazdırılabilir boş sayım formu |
| `Arsiv` | Kapatılan her günün tarihsel satır arşivi |

## Hesaplar

- **Donuk beklenen kapanış** = devreden donuk + gelen − çözündürülen
- **Günlük beklenen kapanış** = devreden yiyecek + dondan çözülen − satış − tadım − atık
- **Fark** = sayılan kapanış − beklenen kapanış (eksi değer kayıptır)
- **UPH (100 adisyon başına satış)** = ortalama günlük satış ÷ referans günlük adisyon × 100
- **Önerilen çözündürme** = `MAX(0; YUKARIYUVARLA(tahmini talep − eldeki günlük beklenen; 0))`
- **SKT** = çözündürme bitiş tarihi + raf ömrü − 1 (mark-out saatinde satıştan kaldırılır)
- **Fire oranı %** = 30 günlük atık ÷ (30 günlük satış + 30 günlük atık)

## Makrolar

| Kısayol | Makro | Ne yapar |
|---|---|---|
| `Ctrl+Shift+S` | `SatislariAl` | ERP satış raporu dosyasını seçtirir, SKU ve adet sütunlarını toplar, `Markout` sayfasına yazar; formda tanımlı olmayan yeni SKU varsa uyarır. |
| `Ctrl+Shift+P` | `GunSonuPDF` | Formu `Markout_<mağaza>_<tarih>.pdf` adıyla çalışma kitabının klasörüne kaydeder. |
| `Ctrl+Shift+K` | `GunuKapat` | Eksik sayım varsa uyarır, günü `Arsiv` sayfasına kaydeder, kapanışları yarının açılış stoğuna ve çözündürülenleri yarının "dondan çözülen" sütununa devredip tarihi ilerletir. |

## Makroları Etkinleştirme ve Deneme

1. `markout-formu.xlsx`, `markout.bas` ve `ornek-satis-raporu.xlsx` dosyalarını aynı klasöre indirin.
2. Formülleri, `SKT`, `Cozundurme_Plani` ve `Fire_Analizi` sayfalarını doğrudan `.xlsx` dosyasında inceleyebilirsiniz.
3. Makroları çalıştırmak için Excel'de **Alt+F11** > **File > Import File…** adımından `markout.bas` dosyasını seçip kitabı `.xlsm` olarak kaydedin (`Alt+F8` > `KisayollariAta` ile kısayolları tanımlayın).
4. `Markout` sayfasında **Ctrl+Shift+S** ile `ornek-satis-raporu.xlsx` dosyasını seçtiğinizde satış sütunu dolar ve örnek rapordaki tanımsız `109999` SKU kodu için kontrol uyarısı üretilir.

## Lisans

MIT. Ayrıntılar için `LICENSE` dosyasına bakın.
