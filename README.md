# Mağaza Gün Sonu (Mark-Out) ve Çözündürme Formu

*Excel + VBA form for end-of-day stock count, sales reconciliation, thaw scheduling and next-day carryover in a coffee store. Documentation in Turkish; all data is synthetic.*

Bir kahve mağazasında donuk ve günlük yiyeceklerin gün sonu sayımını, satışla karşılaştırmasını ve ertesi güne devrini yöneten Excel + VBA çalışma kitabı.

> **Not:** Bu depo, bir kahve zincirinde bölge raporlama sisteminin mağaza ayağı olarak kurduğum ve kullandığım formun **sahte veriyle yeniden kurulmuş sürümüdür**. Orijinal dosyanın yapısı (sütunlar, hesaplar, SKT sayfası, yazdırılabilir form) korunmuştur. Ürünler, SKU'lar ve satış rakamları uydurmadır.
>
> Orijinal sistemde satış verisi şirketin ERP sisteminden (Oracle) otomatik indiriliyordu. O bölüm şirket bilgisayarlarında kaldığı için burada yoktur; yerine, indirilmiş satış raporunu seçip okuyan bir makro vardır.

## Hangi problemi çözüyor

Her akşam kapanışta her ürün için şu soruların cevabı gerekir: Donukta ne kadar kaldı? Gün içinde ne kadar satıldı, ne kadar atıldı? Sayılan stok beklenenle tutuyor mu? Yarın rafa hangi ürün hangi son kullanma tarihiyle çıkacak?

Bu formda:

- Satış adetleri rapordan tek tuşla SKU'ya göre gelir.
- Beklenen kapanış ve fark (kayıp) kendiliğinden hesaplanır.
- Gün kapatıldığında bugünün kapanışı yarının stoğu olur ve form ertesi güne hazır hâle gelir. Günün kaydı arşive yazılır.
- Çözündürmeye alınan ürünlerin son kullanma tarihleri ayrı sayfada hesaplanır.

## Öncesi

Satış adetleri rapordan ürün ürün bakılarak forma elle yazılıyor, kapanış farkları elle hesaplanıyor, ertesi günün devreden stokları elle taşınıyordu.

## Dosyalar

| Dosya | İçerik |
|---|---|
| `markout-formu.xlsx` | Çalışma kitabı (makrosuz). Formüller ve örnek veri içinde |
| `markout.bas` | VBA modülü. Excel'e içe aktarılır |
| `ornek-satis-raporu.xlsx` | Makroyu denemek için sahte günlük satış raporu |

`markout.bas` içinde Türkçe karakter kullanılmamıştır. VBA düzenleyicisi `.bas` dosyalarını sistemin ANSI kodlamasıyla okuduğu için, bu sayede dosya hem GitHub'da hem de içe aktarıldıktan sonra VBA'da doğru görünür.

## Sayfalar

| Sayfa | Ne yapar |
|---|---|
| `Nasil_Kullanilir` | Günlük akış ve hesapların açıklaması |
| `Markout` | Ana form: donuk ve günlük bölüm, beklenen kapanış ve fark |
| `SKT` | Bugün çözündürmeye alınan ve dünden devreden ürünlerin çözündürme saatleri ve son kullanma tarihleri |
| `Bos_Form` | Kapanış sayımında elle doldurmak için yazdırılabilir boş form |
| `Satis_Ortalama` | Haftalık ve aylık satıştan ortalama günlük satış |
| `Ayarlar` | Mağaza adı, ortalama dayanağı, çözündürme saatleri, satış raporunun biçimi |
| `Arsiv` | Kapatılan her günün satırları |

## Hesaplar

- **Donuk beklenen kapanış** = devreden donuk + gelen − çözündürülen
- **Günlük beklenen kapanış** = devreden yiyecek + dondan çözülen − satış − tadım − atık
- **Fark** = sayılan kapanış − beklenen kapanış (eksi değer kayıptır)
- **Dondan çözülen**: bir önceki gün çözündürmeye alınan miktar. Çözündürme 16:00'da başlar, ertesi sabah 08:00'de biter.
- **SKT** = çözündürme bitiş tarihi + raf ömrü − 1. Ürün o gün mark-out saatinde (22:00) satıştan kaldırılır.
- **Ortalama günlük satış** = son 7 gün ÷ 7 veya son 30 gün ÷ 30. Bayram gibi kısa dönemlerde haftalık, genel planlamada aylık kullanılır. Çözündürme miktarı yöneticinin kararıdır; bu sütun karar için yardımcıdır.

## Makrolar

| Kısayol | Makro | Ne yapar |
|---|---|---|
| Ctrl+Shift+S | `SatislariAl` | Satış raporu dosyasını seçtirir. SKU ve adet sütunlarını okur, aynı SKU'nun satırlarını toplar, forma yazar. Raporda olup formda olmayan SKU'ları listeler. |
| Ctrl+Shift+P | `GunSonuPDF` | Formu `Markout_<mağaza>_<tarih>.pdf` olarak çalışma kitabının klasörüne kaydeder. |
| Ctrl+Shift+K | `GunuKapat` | Boş sayım varsa uyarır. Satırları arşive yazar. Sayılan kapanışları yarının devreden stoğuna, çözündürülenleri yarının "dondan çözülen" sütununa taşır. Günlük girişleri temizler, tarihi bir gün ilerletir. |

Formül sütunlarına makro dokunmaz; hesaplar her zaman sayfada görünür kalır.

## Kurulum

1. `markout-formu.xlsx` ve `markout.bas` dosyalarını indirin.
2. Çalışma kitabını açın, **Alt+F11** ile VBA düzenleyicisini açın.
3. **File > Import File…** ile `markout.bas` dosyasını seçin.
4. Düzenleyiciyi kapatın. Dosyayı **Farklı Kaydet > Excel Makro İçerebilen Çalışma Kitabı (\*.xlsm)** olarak kaydedin.
5. Kısayolların dosya her açıldığında hazır olması için: VBA düzenleyicisinde sol paneldeki **ThisWorkbook**'a çift tıklayın ve şu üç satırı yapıştırın:

   ```vb
   Private Sub Workbook_Open()
       KisayollariSessizAta
   End Sub
   ```

   Dosyayı kaydedip kapatın ve yeniden açın. (Bu adım atlanırsa kısayollar için her açılışta **Alt+F8** > `KisayollariAta` çalıştırılır; makrolar Alt+F8 listesinden doğrudan da seçilebilir.)

## Deneme

1. `Markout` sayfasında **Ctrl+Shift+S**'ye basıp `ornek-satis-raporu.xlsx` dosyasını seçin. Satış sütunu dolar; rapordaki `109999` kodlu ürünün formda olmadığı uyarısı çıkar.
2. Birkaç satıra donuk ve günlük "Sayılan Kapanış" girin; fark sütunları hesaplanır.
3. **Ctrl+Shift+K** ile günü kapatın. `Arsiv` sayfasına satırlar eklenir, tarih bir gün ilerler, sayılan kapanışlar devreden sütunlarına geçer.

## Kendi verinize uyarlama

| Nerede | Ne değiştirilir |
|---|---|
| `Markout` | Ürün satırları: kategori başlığı (SKU boş), SKU, ürün adı, raf ömrü (gün) |
| `Satis_Ortalama` | Aynı SKU'larla haftalık ve aylık satış |
| `Ayarlar` B8–B10 | Satış raporunda SKU'nun ve adedin kaçıncı sütunda olduğu, verinin kaçıncı satırda başladığı |
| `Ayarlar` B5–B7 | Çözündürme ve mark-out saatleri |
| `markout.bas` yapılandırma bloğu | Sayfa veya sütun yerleri değişirse sütun numaraları |

Yeni ürün satırı eklendiğinde `SKT` ve `Bos_Form` sayfalarına da aynı satır için formül eklenmelidir.

## Sınırlar

- ERP'den otomatik veri indirme bu sürümde yoktur; rapor elle indirilip seçilir.
- `GunuKapat` geri alınamaz. Kapatmadan önce PDF almak önerilir.
- Makrolar Windows'ta Excel 2010 ve sonrasıyla yazılmıştır. Mac ve LibreOffice'te denenmemiştir.

## Lisans

MIT. Ayrıntılar için `LICENSE` dosyasına bakın.
