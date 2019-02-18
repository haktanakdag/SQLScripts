Public Class Mikro
    Dim voi As New voilib
    Dim dbl As New DBLib
    Function MikroFirmaKontrol(firmakodu As String) As Boolean
        Dim firmalist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM FIRMALAR where fir_SpecRECno =" & firmakodu)
        If firmalist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroFaturaKontrol(belgekodu As String) As Boolean
        Dim belgelist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM CARI_HESAP_HAREKETLERI where cha_belge_no='" & belgekodu & "'")
        If belgelist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroSubeKontrol(subekodu As String)
        Dim subelist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM SUBELER where Sube_DBCno =" & subekodu)
        If subelist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroCariKontrol(carikodu As String)
        Dim carilist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM CARI_HESAPLAR where cari_kod ='" & carikodu & "'")
        If carilist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroDepoKontrol(depokodu As String)
        Dim depolist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM DEPOLAR where dep_no ='" & depokodu & "'")
        If depolist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroPersonelKontrol(stkodu As String)
        Dim stlist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM CARI_PERSONEL_TANIMLARI where cari_per_kod ='" & stkodu & "'")
        If stlist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroKasaKontrol(kasakodu As String)
        Dim stlist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM KASALAR where kas_kod ='" & kasakodu & "'")
        If stlist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroUrunKontrol(urunkodu As String)
        Dim stlist As List(Of DataRow) = voi.QueryDondurMikro("SELECT * FROM STOKLAR where sto_kod ='" & urunkodu & "'")
        If stlist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function MikroBirimKontrol(birimkodu As String)
        Dim stlist As List(Of DataRow) = voi.QueryDondurMikroAyar("SELECT * FROM STOK_BIRIMLERI_CHOOSE_2 where msg_S_0070 ='" & birimkodu & "'")
        If stlist.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function CariHesapBaslikKaydet(Tip As String, firmano As Integer, subeno As Integer, evrakseri As String, evraksirano As Integer, belgetarih As String, belgeno As String, stkodu As String, carikod As String, bruttutar As String, nettutar As String, kdvtutar1 As String, kdvtutar2 As String, kdvtutar3 As String, kdvtutar4 As String, iskontotutar As String, ebelge As String, referansno As String, aciklama As String) As Guid

        Dim voi As New voilib
        Dim querykolonlarCari As String = ""
        Dim querydegerlerCari As String = ""
        Dim insertqueryCari As String = ""
        Dim querylistCari As List(Of DataRow) = voi.QueryDondur("SELECT TIP,KOLONADI, " & Tip & "DEGER FROM TBL_ENT_MIKRO_CARIHESAPKOLON_ESLESTIRME_V16 WHERE DURUM=1")
        Dim querybaslikCari As String = "INSERT INTO CARI_HESAP_HAREKETLERI ( "
        Dim querysonucCari As String = " ) VALUES ("
        For Each q As DataRow In querylistCari
            querykolonlarCari = querykolonlarCari & q("KOLONADI") & ","
        Next
        'querykolonlar = querykolonlar.Substring(0, querykolonlar.Count() - 1)

        querybaslikCari = querybaslikCari & querykolonlarCari

        querybaslikCari = querybaslikCari & "cha_firmano,"
        querybaslikCari = querybaslikCari & "cha_subeno,"
        querybaslikCari = querybaslikCari & "cha_evrak_tip,"
        querybaslikCari = querybaslikCari & "cha_evrakno_seri,"
        querybaslikCari = querybaslikCari & "cha_evrakno_sira,"
        querybaslikCari = querybaslikCari & "cha_tarihi,"
        querybaslikCari = querybaslikCari & "cha_cinsi,"
        querybaslikCari = querybaslikCari & "cha_belge_no,"
        querybaslikCari = querybaslikCari & "cha_belge_tarih,"
        querybaslikCari = querybaslikCari & "cha_satici_kodu,"
        querybaslikCari = querybaslikCari & "cha_kod,"
        querybaslikCari = querybaslikCari & "cha_ciro_cari_kodu,"
        querybaslikCari = querybaslikCari & "cha_d_kur,"
        querybaslikCari = querybaslikCari & "cha_altd_kur,"
        querybaslikCari = querybaslikCari & "cha_meblag,"
        querybaslikCari = querybaslikCari & "cha_aratoplam,"
        querybaslikCari = querybaslikCari & "cha_ft_iskonto1,"
        querybaslikCari = querybaslikCari & "cha_vergi1,"
        querybaslikCari = querybaslikCari & "cha_vergi2,"
        querybaslikCari = querybaslikCari & "cha_vergi3,"
        querybaslikCari = querybaslikCari & "cha_vergi4,"
        querybaslikCari = querybaslikCari & "cha_vade,"
        querybaslikCari = querybaslikCari & "cha_kasa_hizmet,"
        querybaslikCari = querybaslikCari & "cha_kasa_hizkod,"
        querybaslikCari = querybaslikCari & "cha_e_islem_turu,"
        querybaslikCari = querybaslikCari & "cha_specRecNo,"
        querybaslikCari = querybaslikCari & "cha_aciklama,"
        querybaslikCari = querybaslikCari & "cha_Guid"



        For Each q As DataRow In querylistCari
            If q("TIP") = "KARAKTER" Then
                querydegerlerCari = querydegerlerCari & "'" & q(Tip & "DEGER") & "',"
            ElseIf q("TIP") = "TARIH" Then
                If q(Tip & "DEGER").ToString = "GETDATE()" Then
                    querydegerlerCari = querydegerlerCari & "" & q(Tip & "DEGER") & ","
                Else
                    querydegerlerCari = querydegerlerCari & "'" & q(Tip & "DEGER") & "',"
                End If
            Else
                querydegerlerCari = querydegerlerCari & q(Tip & "DEGER") & ","
            End If

        Next
        'querydegerler = querydegerler.Substring(0, querydegerler.Count - 1)

        'querydegerlerCari = querydegerlerCari & "51,"
        querydegerlerCari = querydegerlerCari & firmano & "," 'Firma No
        querydegerlerCari = querydegerlerCari & subeno & "," 'Şube No

        Dim fttip As String = ""
        If Tip = "FATURA" Then
            fttip = "63"
        End If
        If Tip = "NAKITKASA" Then
            fttip = "1"

        End If

        querydegerlerCari = querydegerlerCari & fttip & "," 'Evrak Tip
        querydegerlerCari = querydegerlerCari & "'" & evrakseri & "'," 'cha_evrakno_seri
        querydegerlerCari = querydegerlerCari & "" & evraksirano & "," 'cha_evrakno_sira
        querydegerlerCari = querydegerlerCari & "'" & belgetarih & "'," 'Evrak Tarihi
        Dim cha_cinsi As String = ""
        If Tip = "FATURA" Then
            cha_cinsi = "6"
        End If
        If Tip = "NAKITKASA" Then
            cha_cinsi = "0"

        End If
        querydegerlerCari = querydegerlerCari & cha_cinsi & "," 'cha_cinsi
        querydegerlerCari = querydegerlerCari & "'" & belgeno & "'," 'Belge No
        querydegerlerCari = querydegerlerCari & "'" & belgetarih & "'," 'Belge Tarihi
        querydegerlerCari = querydegerlerCari & "'" & stkodu & "'," 'Satıcı Kodu
        querydegerlerCari = querydegerlerCari & "'" & carikod & "'," 'Cari Kodu
        querydegerlerCari = querydegerlerCari & "'" & carikod & "'," 'Ciro Cari Kodu

        Dim dovizlist As List(Of DataRow) = voilib.QueryDondurMikroAyar("Select TOP 1 IsNull(dov_fiyat1, 1.0) as dovtut from DOVIZ_KURLARI WITH (NOLOCK) WHERE dov_no=1 AND dov_tarih<='" & belgetarih & "' AND dov_fiyat1>0 ORDER BY dov_tarih DESC")
        Dim dovtut As String = ""
        If (dovizlist.Count > 0) Then
            dovtut = dovizlist(0)("dovtut").ToString
        Else
            dovtut = "1"
        End If
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(dovtut).ToString, ",", ".") & "," 'cha_d_kur
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(dovtut).ToString, ",", ".") & "," 'cha_altd_kur
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(bruttutar).ToString, ",", ".") & "," 'cha_meblag
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(nettutar).ToString, ",", ".") & "," 'cha_aratoplam
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(iskontotutar).ToString, ",", ".") & "," 'cha_ft_iskonto1 '
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(kdvtutar1).ToString, ",", ".") & "," 'cha_vergi1
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(kdvtutar2).ToString, ",", ".") & "," 'cha_vergi2
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(kdvtutar3).ToString, ",", ".") & "," 'cha_vergi3
        querydegerlerCari = querydegerlerCari & Replace(Convert.ToDouble(kdvtutar4).ToString, ",", ".") & "," 'cha_vergi4

        If Tip = "FATURA" Then
            querydegerlerCari = querydegerlerCari & "0," 'cha_vade
        End If
        If Tip = "NAKITKASA" Then
            querydegerlerCari = querydegerlerCari & belgetarih & "," 'cha_vade
        End If

        If Tip = "FATURA" Then
            querydegerlerCari = querydegerlerCari & "0," 'cha_kasa_hizmet
        End If
        If Tip = "NAKITKASA" Then
            querydegerlerCari = querydegerlerCari & "4," 'cha_kasa_hizmet
        End If

        If Tip = "FATURA" Then
            querydegerlerCari = querydegerlerCari & "''," 'cha_kasa_hizkod
        End If
        If Tip = "NAKITKASA" Then
            querydegerlerCari = querydegerlerCari & "'100.01.001'" 'cha_kasa_hizkod
        End If

        If Tip = "FATURA" Then
            querydegerlerCari = querydegerlerCari & "" & ebelge & ","
            querydegerlerCari = querydegerlerCari & "'" & referansno & "'," 'cha_specRecNo
            querydegerlerCari = querydegerlerCari & "'" & aciklama & "'," 'cha_aciklama
        End If

        Dim dbl As New DBLib

        Dim sqlUnique As String = "INSERT INTO UNIQUE_TABLE (TIP,UID) VALUES('" & Tip & "',NEWID())"
        Dim carihareketid As Guid = dbl.Unique_Kaydet_Cek(sqlUnique, Tip)

        querydegerlerCari = querydegerlerCari & "'" & carihareketid.ToString & "'"


        querydegerlerCari = querydegerlerCari & ")"
        querysonucCari = querysonucCari & querydegerlerCari
        insertqueryCari = querybaslikCari & querysonucCari
        Try

            dbl.CariHesapKaydet(insertqueryCari)
            dbl.CariHareketKodUpdateFatura(carihareketid)
            dbl.Unique_Bosalt("DELETE FROM UNIQUE_TABLE WHERE TIP='" & Tip & "'")
        Catch ex As Exception
            voi.Entlogyaz("", belgeno, "Başlık bilgileri kayıt sırasında hata -> " & ex.Message.ToString, 0)

        End Try
        Return carihareketid
    End Function

    Function SatisFaturaDetayKaydet(firmano As Integer, subeno As Integer, carihareketid As Guid, belgetarih As String, belgeno As String, evraksira As Integer, kalemsira As Integer, stokkodu As String, carikodu As String, stkodu As String, miktar As Double, birim As String, tutar As Double, vergi As Double, kdvoransira As Integer, fataltisktutar As Double, girisdepo As Integer, cikisdepo As Integer, isktutar1 As Double, isktutar2 As Double, isktutar3 As Double, isktutar4 As Double, isktutar5 As Double, isktutar6 As Double)

        Dim dbl As New DBLib
        Dim voi As New voilib

        Dim querykolonlarStok As String = ""
        Dim querydegerlerStok As String = ""
        Dim insertqueryStok As String = ""
        Dim querylistStok As List(Of DataRow) = voi.QueryDondur("Select TIP, KOLONADI, DEGER FROM TBL_ENT_MIKRO_STOKHAREKETKOLON_ESLESTIRME_V16")
        Dim querybaslikStok As String = "INSERT INTO dbo.STOK_HAREKETLERI  ( "
        Dim querysonucStok As String = " ) VALUES ("
        For Each q As DataRow In querylistStok
            querykolonlarStok = querykolonlarStok & q("KOLONADI") & ","
        Next
        'querykolonlar = querykolonlar.Substring(0, querykolonlar.Count() - 1)

        querybaslikStok = querybaslikStok & querykolonlarStok

        'querybaslikStok = querybaslikStok & "sth_create_date,"
        'querybaslikStok = querybaslikStok & "sth_lastup_date,"
        querybaslikStok = querybaslikStok & "sth_firmano,"
        querybaslikStok = querybaslikStok & "sth_subeno,"
        querybaslikStok = querybaslikStok & "sth_tarih,"
        querybaslikStok = querybaslikStok & "sth_tip,"
        querybaslikStok = querybaslikStok & "sth_evraktip,"
        querybaslikStok = querybaslikStok & "sth_evrakno_seri,"
        querybaslikStok = querybaslikStok & "sth_evrakno_sira,"
        querybaslikStok = querybaslikStok & "sth_belge_no,"
        querybaslikStok = querybaslikStok & "sth_belge_tarih,"
        querybaslikStok = querybaslikStok & "sth_stok_kod,"
        querybaslikStok = querybaslikStok & "sth_cari_kodu,"
        querybaslikStok = querybaslikStok & "sth_plasiyer_kodu,"
        querybaslikStok = querybaslikStok & "sth_har_doviz_kuru,"
        querybaslikStok = querybaslikStok & "sth_alt_doviz_kuru,"
        querybaslikStok = querybaslikStok & "sth_stok_doviz_kuru,"
        querybaslikStok = querybaslikStok & "sth_miktar,"
        querybaslikStok = querybaslikStok & "sth_miktar2,"
        querybaslikStok = querybaslikStok & "sth_birim_pntr,"
        querybaslikStok = querybaslikStok & "sth_tutar,"
        querybaslikStok = querybaslikStok & "sth_vergi_pntr,"
        querybaslikStok = querybaslikStok & "sth_vergi,"
        querybaslikStok = querybaslikStok & "sth_iskonto1,"
        querybaslikStok = querybaslikStok & "sth_iskonto2,"
        querybaslikStok = querybaslikStok & "sth_iskonto3,"
        querybaslikStok = querybaslikStok & "sth_iskonto4,"
        querybaslikStok = querybaslikStok & "sth_iskonto5,"
        querybaslikStok = querybaslikStok & "sth_iskonto6,"
        ' querybaslikStok = querybaslikStok & "sth_fat_recid_dbcno,"
        ' querybaslikStok = querybaslikStok & "sth_fat_recid_recno,"
        querybaslikStok = querybaslikStok & "sth_giris_depo_no,"
        querybaslikStok = querybaslikStok & "sth_cikis_depo_no,"
        querybaslikStok = querybaslikStok & "sth_malkbl_sevk_tarihi,"
        querybaslikStok = querybaslikStok & "sth_Guid"
        'querybaslikStok = querybaslikStok & "sth_adres_no"

        For Each q As DataRow In querylistStok
            If q("TIP") = "KARAKTER" Then
                querydegerlerStok = querydegerlerStok & "'" & q("DEGER") & "',"
            ElseIf q("TIP") = "TARIH" Then
                If q("DEGER") = "GETDATE()" Then
                    querydegerlerStok = querydegerlerStok & "" & q("DEGER") & ","
                Else
                    querydegerlerStok = querydegerlerStok & "'" & q("DEGER") & "',"
                End If
            Else
                querydegerlerStok = querydegerlerStok & q("DEGER") & ","
            End If

        Next
        'querydegerler = querydegerler.Substring(0, querydegerler.Count - 1)

        'querydegerlerStok = querydegerlerStok & "GETDATE(),"
        'querydegerlerStok = querydegerlerStok & "GETDATE(),"
        querydegerlerStok = querydegerlerStok & firmano & ","
        querydegerlerStok = querydegerlerStok & subeno & ","
        querydegerlerStok = querydegerlerStok & "'" & belgetarih & "'," 'sth_tarih
        querydegerlerStok = querydegerlerStok & "1," 'sth_tip
        querydegerlerStok = querydegerlerStok & "4," 'sth_evraktip
        querydegerlerStok = querydegerlerStok & "''," 'sth_evrakno_seri,
        querydegerlerStok = querydegerlerStok & "" & evraksira & "," 'sth_evrakno_sira
        querydegerlerStok = querydegerlerStok & "'" & belgeno & "'," 'Belge No
        querydegerlerStok = querydegerlerStok & "'" & belgetarih & "'," 'Belge Tarihi
        querydegerlerStok = querydegerlerStok & "'" & stokkodu & "'," 'Stok Kodu
        querydegerlerStok = querydegerlerStok & "'" & carikodu & "'," 'Cari Kodu
        querydegerlerStok = querydegerlerStok & "'" & stkodu & "'," 'sth_plasiyer_kodu
        querydegerlerStok = querydegerlerStok & "1," 'sth_har_doviz_kuru
        querydegerlerStok = querydegerlerStok & "1," 'sth_alt_doviz_kuru
        querydegerlerStok = querydegerlerStok & "1," 'sth_stok_doviz_kuru
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(miktar).ToString, ",", ".") & "," 'sth_miktar
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(miktar).ToString, ",", ".") & "," 'sth_miktar2
        querydegerlerStok = querydegerlerStok & "" & birim & "," 'sth_birim_pntr
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(tutar).ToString, ",", ".") & "," 'sth_tutar
        querydegerlerStok = querydegerlerStok & "'" & kdvoransira & "'," 'sth_vergi_pntr
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(vergi).ToString, ",", ".") & "," 'sth_vergi
        'querydegerlerStok = querydegerlerStok & "2," 'sth_fat_recid_recno
        If isktutar1 > 0 Then
            querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(isktutar1).ToString, ",", ".") & "," 'sth_iskonto1
        Else
            querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(fataltisktutar).ToString, ",", ".") & "," 'sth_iskonto1
        End If
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(isktutar2).ToString, ",", ".") & "," 'sth_iskonto2
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(isktutar3).ToString, ",", ".") & "," 'sth_iskonto3
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(isktutar4).ToString, ",", ".") & "," 'sth_iskonto4
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(isktutar5).ToString, ",", ".") & "," 'sth_iskonto5
        querydegerlerStok = querydegerlerStok & "" & Replace(Convert.ToDouble(isktutar6).ToString, ",", ".") & "," 'sth_iskonto6
        querydegerlerStok = querydegerlerStok & "" & girisdepo & "," 'sth_giris_depo_no
        querydegerlerStok = querydegerlerStok & "" & cikisdepo & "," 'sth_cikis_depo_no
        querydegerlerStok = querydegerlerStok & "'" & belgetarih & "'," 'sth_malkbl_sevk_tarihi
        'querydegerlerStok = querydegerlerStok & "1" 'sth_adres_no



        Dim sqlUnique As String = "INSERT INTO UNIQUE_TABLE (TIP,UID) VALUES('Stok',NEWID())"
        Dim stokhareketid As Guid = dbl.Unique_Kaydet_Cek(sqlUnique, "Stok")

        querydegerlerStok = querydegerlerStok & "'" & stokhareketid.ToString & "'"  'cha_altd_kur

        querydegerlerStok = querydegerlerStok & ")"

        querysonucStok = querysonucStok & querydegerlerStok

        insertqueryStok = querybaslikStok & querysonucStok

        Try
            dbl.StokHareketKaydet(insertqueryStok)
            dbl.StokHareketUpdateFatura(stokhareketid, carihareketid)
            dbl.Unique_Bosalt("DELETE FROM UNIQUE_TABLE WHERE TIP='Stok'")
        Catch ex As Exception
            voi.Entlogyaz("", belgeno, "Detay bilgileri kayıt sırasında hata -> " & ex.Message.ToString, 0)
        End Try
    End Function
End Class
