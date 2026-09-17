Attribute VB_Name = "Markout"
'===============================================================================
' MAGAZA GUN SONU (MARK-OUT) FORMU - MAKROLAR
'
' Bu modul uc isi yapar:
'   SatislariAl   (Ctrl+Shift+S) : Satis raporu dosyasini sectirir, SKU'ya gore
'                                  satis adetlerini Markout sayfasina yazar.
'   GunSonuPDF    (Ctrl+Shift+P) : Markout sayfasini tarihli PDF olarak kaydeder.
'   GunuKapat     (Ctrl+Shift+K) : Gunu arsivler, kapanislari yarinin stogu yapar,
'                                  gunluk girisleri temizler, tarihi ilerletir.
'
' Kurulum: Alt+F11 > File > Import File > markout.bas
'          Kisayollarin her acilista hazir olmasi icin ThisWorkbook'a
'          Workbook_Open eklenir (README > Kurulum, 5. adim).
'===============================================================================
Option Explicit

'------------------------------ YAPILANDIRMA ----------------------------------
' Sayfa adlari
Private Const SAYFA_FORM As String = "Markout"
Private Const SAYFA_AYAR As String = "Ayarlar"
Private Const SAYFA_ARSIV As String = "Arsiv"

' Markout sayfasinda urun satirlarinin basladigi satir
Private Const ILK_SATIR As Long = 4

' Tarih hucresi
Private Const HUCRE_TARIH As String = "F1"

' Markout sutun numaralari (A=1, B=2, ...)
Private Const S_SKU As Long = 2            ' B  SKU
Private Const S_URUN As Long = 3           ' C  Urun adi
Private Const S_DEV_DONUK As Long = 6      ' F  Devreden donuk
Private Const S_GELEN As Long = 7          ' G  Gelen
Private Const S_COZDURULEN As Long = 8     ' H  Cozundurulen (bugun donuktan alinan)
Private Const S_DONUK_BEK As Long = 9      ' I  Donuk beklenen kapanis (formul)
Private Const S_DONUK_SAY As Long = 10     ' J  Donuk sayilan kapanis
Private Const S_DONUK_FARK As Long = 11    ' K  Donuk fark (formul)
Private Const S_DEV_YIYECEK As Long = 13   ' M  Devreden yiyecek
Private Const S_COZULEN As Long = 14       ' N  Dondan cozulen (dun cozundurulen)
Private Const S_SATIS As Long = 15         ' O  Satis
Private Const S_TADIM As Long = 16         ' P  Tadim
Private Const S_ATIK As Long = 17          ' Q  Atik
Private Const S_GUN_BEK As Long = 18       ' R  Gunluk beklenen kapanis (formul)
Private Const S_GUN_SAY As Long = 19       ' S  Gunluk sayilan kapanis
Private Const S_GUN_FARK As Long = 20      ' T  Gunluk fark (formul)
Private Const S_BITIS As Long = 21         ' U  Urun bitis saati

' Ayarlar sayfasinda satis raporu bicimini tutan hucreler
Private Const AYAR_SKU_SUTUN As String = "B8"
Private Const AYAR_ADET_SUTUN As String = "B9"
Private Const AYAR_ILK_VERI As String = "B10"
'-------------------------------------------------------------------------------


'===============================================================================
' Kisayollari atar. Dosya ilk kez .xlsm olarak kaydedildikten sonra bir kez
' calistirilir. (Excel'de "+" Shift, "^" Ctrl demektir.)
'===============================================================================
Public Sub KisayollariAta()
    KisayollariSessizAta
    MsgBox "Kisayollar atandi:" & vbCrLf & _
           "Ctrl+Shift+S  Satislari al" & vbCrLf & _
           "Ctrl+Shift+P  Gun sonu PDF" & vbCrLf & _
           "Ctrl+Shift+K  Gunu kapat", vbInformation
End Sub


'===============================================================================
' Kisayollari mesaj gostermeden atar. ThisWorkbook icindeki Workbook_Open
' bunu cagirirsa kisayollar dosya her acildiginda kendiliginden hazir olur.
'===============================================================================
Public Sub KisayollariSessizAta()
    Application.OnKey "^+s", "SatislariAl"
    Application.OnKey "^+p", "GunSonuPDF"
    Application.OnKey "^+k", "GunuKapat"
End Sub


'===============================================================================
' SATISLARI AL
' 1. Kullaniciya satis raporu dosyasini sectirir.
' 2. Dosyayi salt okunur acar, SKU ve adet sutunlarini okur.
'    Ayni SKU raporda birden fazla satirdaysa adetler toplanir.
' 3. Markout'taki her urun satirina kendi satis adedini yazar
'    (raporda olmayan urune 0).
' 4. Raporda olup formda olmayan SKU'lari sonunda listeler.
'===============================================================================
Public Sub SatislariAl()
    Dim form As Worksheet, ayar As Worksheet
    Dim dosyaYolu As Variant
    Dim rapor As Workbook, raporSayfa As Worksheet
    Dim satislar As Object, eslesen As Object
    Dim skuSutun As Long, adetSutun As Long, ilkVeri As Long
    Dim r As Long, sonSatir As Long
    Dim sku As String, adet As Double
    Dim yazilan As Long, bulunamayan As String
    Dim anahtar As Variant

    Set form = ThisWorkbook.Worksheets(SAYFA_FORM)
    Set ayar = ThisWorkbook.Worksheets(SAYFA_AYAR)
    skuSutun = ayar.Range(AYAR_SKU_SUTUN).Value
    adetSutun = ayar.Range(AYAR_ADET_SUTUN).Value
    ilkVeri = ayar.Range(AYAR_ILK_VERI).Value

    ' 1. Dosya secimi (vazgecilirse False doner)
    dosyaYolu = Application.GetOpenFilename( _
        FileFilter:="Satis raporu (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", _
        Title:="Gunluk satis raporunu secin")
    If VarType(dosyaYolu) = vbBoolean Then Exit Sub

    On Error GoTo Hata
    Application.ScreenUpdating = False

    ' 2. Raporu oku -> sozluk: SKU => toplam adet
    Set satislar = CreateObject("Scripting.Dictionary")
    Set rapor = Workbooks.Open(Filename:=dosyaYolu, ReadOnly:=True)
    Set raporSayfa = rapor.Worksheets(1)
    sonSatir = raporSayfa.Cells(raporSayfa.Rows.Count, skuSutun).End(xlUp).Row

    For r = ilkVeri To sonSatir
        sku = SkuTemizle(raporSayfa.Cells(r, skuSutun).Value)
        If sku <> "" And IsNumeric(raporSayfa.Cells(r, adetSutun).Value) Then
            adet = CDbl(raporSayfa.Cells(r, adetSutun).Value)
            satislar(sku) = satislar(sku) + adet
        End If
    Next r
    rapor.Close SaveChanges:=False
    Set rapor = Nothing

    ' 3. Forma yaz
    Set eslesen = CreateObject("Scripting.Dictionary")
    sonSatir = SonUrunSatiri(form)
    For r = ILK_SATIR To sonSatir
        sku = SkuTemizle(form.Cells(r, S_SKU).Value)
        If sku <> "" Then
            If satislar.Exists(sku) Then
                form.Cells(r, S_SATIS).Value = satislar(sku)
                eslesen(sku) = True
                yazilan = yazilan + 1
            Else
                form.Cells(r, S_SATIS).Value = 0
            End If
        End If
    Next r

    ' 4. Formda karsiligi olmayan SKU'lar
    For Each anahtar In satislar.Keys
        If Not eslesen.Exists(anahtar) Then
            bulunamayan = bulunamayan & vbCrLf & "  " & anahtar & "  (" & satislar(anahtar) & " adet)"
        End If
    Next anahtar

    Application.ScreenUpdating = True

    If bulunamayan = "" Then
        MsgBox yazilan & " urunun satisi yazildi.", vbInformation
    Else
        MsgBox yazilan & " urunun satisi yazildi." & vbCrLf & vbCrLf & _
               "Raporda olup formda olmayan SKU'lar:" & bulunamayan & vbCrLf & vbCrLf & _
               "Bu urunler forma eklenmeli ya da SKU'lari kontrol edilmeli.", vbExclamation
    End If
    Exit Sub

Hata:
    Application.ScreenUpdating = True
    If Not rapor Is Nothing Then rapor.Close SaveChanges:=False
    MsgBox "Rapor okunamadi: " & Err.Description & vbCrLf & _
           "Ayarlar sayfasindaki SKU / adet sutun numaralarini kontrol edin.", vbCritical
End Sub


'===============================================================================
' GUN SONU PDF
' Markout sayfasini, calisma kitabinin bulundugu klasore
' "Markout_<magaza>_<yyyy-aa-gg>.pdf" adiyla kaydeder.
'===============================================================================
Public Sub GunSonuPDF()
    Dim form As Worksheet
    Dim dosyaAdi As String

    If ThisWorkbook.Path = "" Then
        MsgBox "Once calisma kitabini kaydedin; PDF ayni klasore yazilir.", vbExclamation
        Exit Sub
    End If

    Set form = ThisWorkbook.Worksheets(SAYFA_FORM)
    dosyaAdi = ThisWorkbook.Path & Application.PathSeparator & "Markout_" & _
               DosyaAdinaUygun(CStr(ThisWorkbook.Worksheets(SAYFA_AYAR).Range("B3").Value)) & "_" & _
               Format(form.Range(HUCRE_TARIH).Value, "yyyy-mm-dd") & ".pdf"

    form.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dosyaAdi, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "PDF kaydedildi:" & vbCrLf & dosyaAdi, vbInformation
End Sub


'===============================================================================
' GUNU KAPAT
' 1. Sayilan kapanisi bos satir varsa uyarir.
' 2. Gunun tum urun satirlarini Arsiv sayfasina ekler (deger olarak).
' 3. Yarina devir:
'      devreden donuk   <- bugunun donuk sayilan kapanisi
'      devreden yiyecek <- bugunun gunluk sayilan kapanisi
'      dondan cozulen   <- bugun cozundurmeye alinan miktar
' 4. Gunluk girisleri temizler, tarihi bir gun ilerletir.
' Formul sutunlarina (beklenen kapanis, fark) dokunmaz.
'===============================================================================
Public Sub GunuKapat()
    Dim form As Worksheet, arsiv As Worksheet
    Dim r As Long, sonSatir As Long, hedef As Long
    Dim eksik As Long
    Dim cevap As VbMsgBoxResult
    Dim tarih As Date

    Set form = ThisWorkbook.Worksheets(SAYFA_FORM)
    Set arsiv = ThisWorkbook.Worksheets(SAYFA_ARSIV)
    sonSatir = SonUrunSatiri(form)
    tarih = form.Range(HUCRE_TARIH).Value

    ' 1. Eksik sayim kontrolu
    For r = ILK_SATIR To sonSatir
        If UrunSatiriMi(form, r) Then
            If IsEmpty(form.Cells(r, S_DONUK_SAY).Value) Or IsEmpty(form.Cells(r, S_GUN_SAY).Value) Then
                eksik = eksik + 1
            End If
        End If
    Next r

    If eksik > 0 Then
        cevap = MsgBox(eksik & " urunde sayilan kapanis bos." & vbCrLf & _
                       "Bos kapanis yarina 0 olarak devreder. Yine de kapatilsin mi?", _
                       vbYesNo + vbExclamation, "Eksik sayim")
    Else
        cevap = MsgBox(Format(tarih, "dd.mm.yyyy") & " gunu kapatilsin mi?" & vbCrLf & _
                       "Bu islem geri alinamaz; once PDF almaniz onerilir.", _
                       vbYesNo + vbQuestion, "Gunu kapat")
    End If
    If cevap <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculate

    ' 2. Arsive yaz
    hedef = arsiv.Cells(arsiv.Rows.Count, 1).End(xlUp).Row + 1
    For r = ILK_SATIR To sonSatir
        If UrunSatiriMi(form, r) Then
            arsiv.Cells(hedef, 1).Value = tarih
            arsiv.Cells(hedef, 2).Value = form.Cells(r, S_SKU).Value
            arsiv.Cells(hedef, 3).Value = form.Cells(r, S_URUN).Value
            arsiv.Cells(hedef, 4).Value = form.Cells(r, S_DONUK_BEK).Value
            arsiv.Cells(hedef, 5).Value = form.Cells(r, S_DONUK_SAY).Value
            arsiv.Cells(hedef, 6).Value = form.Cells(r, S_DONUK_FARK).Value
            arsiv.Cells(hedef, 7).Value = form.Cells(r, S_COZDURULEN).Value
            arsiv.Cells(hedef, 8).Value = form.Cells(r, S_SATIS).Value
            arsiv.Cells(hedef, 9).Value = form.Cells(r, S_TADIM).Value
            arsiv.Cells(hedef, 10).Value = form.Cells(r, S_ATIK).Value
            arsiv.Cells(hedef, 11).Value = form.Cells(r, S_GUN_BEK).Value
            arsiv.Cells(hedef, 12).Value = form.Cells(r, S_GUN_SAY).Value
            arsiv.Cells(hedef, 13).Value = form.Cells(r, S_GUN_FARK).Value
            arsiv.Cells(hedef, 14).Value = form.Cells(r, S_BITIS).Value
            arsiv.Cells(hedef, 1).NumberFormat = "dd.mm.yyyy"
            arsiv.Cells(hedef, 14).NumberFormat = "hh:mm"
            hedef = hedef + 1
        End If
    Next r

    ' 3-4. Devir ve temizlik
    For r = ILK_SATIR To sonSatir
        If UrunSatiriMi(form, r) Then
            form.Cells(r, S_DEV_DONUK).Value = Sayi(form.Cells(r, S_DONUK_SAY).Value)
            form.Cells(r, S_DEV_YIYECEK).Value = Sayi(form.Cells(r, S_GUN_SAY).Value)
            form.Cells(r, S_COZULEN).Value = Sayi(form.Cells(r, S_COZDURULEN).Value)

            form.Cells(r, S_GELEN).ClearContents
            form.Cells(r, S_COZDURULEN).ClearContents
            form.Cells(r, S_DONUK_SAY).ClearContents
            form.Cells(r, S_SATIS).ClearContents
            form.Cells(r, S_TADIM).ClearContents
            form.Cells(r, S_ATIK).ClearContents
            form.Cells(r, S_GUN_SAY).ClearContents
            form.Cells(r, S_BITIS).ClearContents
        End If
    Next r

    form.Range(HUCRE_TARIH).Value = tarih + 1
    Application.ScreenUpdating = True

    MsgBox "Gun kapatildi. Yeni tarih: " & Format(tarih + 1, "dd.mm.yyyy"), vbInformation
End Sub


'------------------------------ YARDIMCILAR -----------------------------------

' Satirda SKU varsa urun satiridir; kategori baslik satirlarinda SKU bostur.
Private Function UrunSatiriMi(ws As Worksheet, r As Long) As Boolean
    UrunSatiriMi = (SkuTemizle(ws.Cells(r, S_SKU).Value) <> "")
End Function

' Urun adi sutunundaki son dolu satir.
Private Function SonUrunSatiri(ws As Worksheet) As Long
    SonUrunSatiri = ws.Cells(ws.Rows.Count, S_URUN).End(xlUp).Row
End Function

' SKU'yu karsilastirilabilir metne cevirir:
' sayi olarak gelen 100101 ile metin olarak gelen "100101 " ayni sayilir.
' Bolunmez bosluk (Chr 160) ve bas/son bosluklar atilir.
Private Function SkuTemizle(deger As Variant) As String
    Dim s As String
    If IsError(deger) Or IsEmpty(deger) Then Exit Function
    s = Trim$(Replace(CStr(deger), Chr(160), " "))
    If s <> "" And IsNumeric(s) Then
        If CDbl(s) = Int(CDbl(s)) Then s = Format(CDbl(s), "0")
    End If
    SkuTemizle = s
End Function

' Bos veya sayi olmayan hucreyi 0 kabul eder.
Private Function Sayi(deger As Variant) As Double
    If IsError(deger) Then Exit Function
    If IsNumeric(deger) And Not IsEmpty(deger) Then Sayi = CDbl(deger)
End Function

' Dosya adinda kullanilamayan karakterleri "_" yapar.
Private Function DosyaAdinaUygun(s As String) As String
    Dim yasak As Variant, k As Variant
    yasak = Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
    For Each k In yasak
        s = Replace(s, k, "_")
    Next k
    DosyaAdinaUygun = s
End Function
