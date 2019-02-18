Imports System
Imports System.Configuration
Imports System.Configuration.Install
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.IO
Imports System.Data.OleDb
Imports System.Collections
Imports System.Timers
Imports System.IO.Compression
Imports System.Xml
Imports System.Web
Imports System.ComponentModel
Imports System.Collections.Generic
Imports NetOpenX50
Imports NetsisConnect
Imports System.Runtime.InteropServices
Public Class frmFatura
    Dim lb As New voilib

    Private Sub frmFatura_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim subelerliste As New DBLib
        cmbListe.DataSource = subelerliste.subeleri_getir()
        cmbListe.ValueMember = "LNGPANKOD"
        cmbListe.DisplayMember = "TXTPANSUBEAD"
        txtBitTarih.Text = Now
        txtBasTarih.Text = Now
        rdSecilen.Checked = True
        cmbTip.SelectedIndex = 0
    End Sub

    Private Sub btnPanCek_Click(sender As Object, e As EventArgs) Handles btnPanCek.Click
        Dim lb As New voilib

        Dim kriterstring As String = ""
        Dim veri As New DBLib
        'Dim voi As New library
        Using connection As New SqlConnection(veri.gGetConnectionStringPan())
            'connection.Open()
            'connection.Close()
            Dim subekodu As Integer = cmbListe.SelectedValue
            Dim turstring As String = cmbTip.SelectedIndex.ToString
            Dim bastarih As String = lb.tarihDondurPan(txtBasTarih.Text)
            Dim bittarih As String = lb.tarihDondurPan(txtBitTarih.Text)
            Dim strSQL As String = "EXEC SSP_ENT_FATURA_BASLIK " & subekodu & "," & turstring & ",'" & bastarih & "','" & bittarih & "'"

            Dim da As New SqlDataAdapter(strSQL, connection)
            'Dim da As New SqlDataAdapter(strSQL & " " & kriterstring, connection)
            Dim ds As New DataSet
            da.Fill(ds)
            grdListe.DataSource = ds.Tables(0)
            grdListe.AutoGenerateColumns = True
            grdListe.Refresh()
        End Using
    End Sub

    Private Sub btn_mikroyaaktar_Click(sender As Object, e As EventArgs) Handles btn_mikroyaaktar.Click

        Dim belgekod As String = ""
        If grdListe.Rows.Count = 0 Then
            MessageBox.Show("Aktarılacak Herhangi bir kayıt yok!")
        Else
            If grdListe.Rows(0).Cells(1).Value = "" Then
                MessageBox.Show("Aktarılacak Herhangi bir kayıt yok!")
            Else
                Try

                    If rdSecilen.Checked = True Then
                        Dim say As Integer
                        For say = 0 To grdListe.SelectedRows.Count - 1
                            If grdListe.SelectedRows(say).Cells(4).Value.ToString().Count > 0 Then
                                If grdListe.SelectedRows(say).Index >= 0 Then
                                    AktarimYap(grdListe.SelectedRows(say).Index)
                                End If
                                belgekod = belgekod & "," & grdListe.Rows(say).Cells(16).Value.ToString()

                            Else
                                lb.Entlogyaz(grdListe.SelectedRows(say).Cells(28).Value.ToString(), "", "Fatura aktarımı sırasında hata->." & "Belge numarası boş olamaz!", 0)
                            End If
                        Next
                        btnPanCek.PerformClick()
                        If (belgekod <> "") Then
                            frmLogs.belgekod = belgekod
                            frmLogs.MdiParent = MainForm
                            frmLogs.Show()
                        End If
                    End If
                    If rdTumu.Checked = True Then
                        For tsay As Integer = 0 To grdListe.Rows.Count - 1
                            If grdListe.Rows(tsay).Cells(4).Value.ToString().Count > 0 Then
                                If grdListe.Rows(tsay).Index >= 0 Then
                                    AktarimYap(grdListe.Rows(tsay).Index)
                                End If
                                belgekod = belgekod & "," & grdListe.Rows(tsay).Cells(28).Value.ToString()
                            Else
                                lb.Entlogyaz(grdListe.SelectedRows(tsay).Cells(2).Value.ToString(), "", "Fatura aktarımı sırasında hata->." & "Belge numarası boş olamaz!", 0)
                            End If
                        Next
                        btnPanCek.PerformClick()
                        If (belgekod <> "") Then
                            frmLogs.belgekod = belgekod
                            frmLogs.MdiParent = MainForm
                            frmLogs.Show()
                        End If
                    End If
                Catch ex As Exception
                    lb.Entlogyaz("", "", "Fatura aktarımı sırasında hata->." & ex.Message.ToString, 0)
                    btnPanCek.PerformClick()
                    If (belgekod <> "") Then
                        frmLogs.belgekod = belgekod
                        frmLogs.MdiParent = MainForm
                        frmLogs.Show()
                    End If
                Finally
                    btnPanCek.PerformClick()
                    If (belgekod <> "") Then
                        frmLogs.belgekod = belgekod
                        frmLogs.MdiParent = MainForm
                        frmLogs.Show()
                    End If
                End Try
            End If
        End If
    End Sub

    Function AktarimYap(satirno As Integer)

        'Dim voi As New voilib
        'Dim dbl As New DBLib
        Dim Pan As New Pan

        Dim kaynakbelgeno As String = grdListe.Rows(satirno).Cells(0).Value.ToString()
        Dim hedefbelgeno As String = grdListe.Rows(satirno).Cells(16).Value.ToString()
        Dim mikro As New Mikro

        Dim hatavar As Boolean = False
        Dim firmano As Integer = 0
        Dim subeno As Integer = grdListe.Rows(satirno).Cells(4).Value.ToString()
        Dim carihareketid As Guid

        Dim belgeno As String = grdListe.Rows(satirno).Cells(16).Value.ToString()
        Dim stkodu As String = grdListe.Rows(satirno).Cells(15).Value.ToString()
        Dim carikodu As String = grdListe.Rows(satirno).Cells(9).Value.ToString()
        Dim depokodu As String = grdListe.Rows(satirno).Cells(11).Value.ToString()
        Dim belgetarih As String = grdListe.Rows(satirno).Cells(17).Value.ToString()

        belgetarih = lb.tarihYYYYMMDDondur(belgetarih)

        If mikro.MikroFaturaKontrol(hedefbelgeno) = True Then
            hatavar = True
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, hedefbelgeno & "->Belge Kodu zaten var!", 0)
        End If

        If mikro.MikroFirmaKontrol(0) = False Then
            hatavar = True
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Firma Kodu Bulunamadı", 0)
        End If
        If mikro.MikroSubeKontrol(subeno) = False Then
            hatavar = True
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Şube Kodu Bulunamadı->"& subeno, 0)
        End If
        If mikro.MikroCariKontrol(carikodu) = False Then
            hatavar = True
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Cari Kodu Bulunamadı->" & carikodu, 0)
        End If
        If mikro.MikroPersonelKontrol(stkodu) = False Then
            hatavar = True
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "ST Kodu Bulunamadı->" & stkodu, 0)
        End If
        If mikro.MikroDepoKontrol(depokodu) = False Then
            hatavar = True
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Depo Kodu Bulunamadı->" & depokodu, 0)
        End If


        Dim bruttutar As Double = grdListe.Rows(satirno).Cells(20).Value.ToString()
        Dim nettutar As Double = grdListe.Rows(satirno).Cells(24).Value.ToString()

        Dim kdvtutar1 As Double = grdListe.Rows(satirno).Cells(31).Value.ToString()
        Dim kdvtutar2 As Double = grdListe.Rows(satirno).Cells(32).Value.ToString()
        Dim kdvtutar3 As Double = grdListe.Rows(satirno).Cells(33).Value.ToString()
        Dim kdvtutar4 As Double = grdListe.Rows(satirno).Cells(34).Value.ToString()

        Dim iskontotutar As Double = grdListe.Rows(satirno).Cells(21).Value.ToString()
        Dim ebelge As String = grdListe.Rows(satirno).Cells(29).Value.ToString()
        Dim aciklama As String = grdListe.Rows(satirno).Cells(18).Value.ToString()
        Dim evrakseri As String = grdListe.Rows(satirno).Cells(30).Value.ToString()


        Dim evraklist As List(Of DataRow) = lb.QueryDondurMikro("SELECT MAX(cha_evrakno_sira) evraksirano FROM CARI_HESAP_HAREKETLERI WHERE cha_evrak_tip=63 and cha_evrakno_seri='" & evrakseri & "' and  cha_subeno=" & subeno & " and cha_firmano=" & firmano)

        'Birim Sırayı Öğren
        Dim evraksira As Integer = 0
        If IsDBNull(evraklist(0)("evraksirano")) = False Then
            evraksira = evraklist(0)("evraksirano")
        End If


        Dim lngbelgekod As Integer = kaynakbelgeno

        Dim detayquery As String = "EXEC SSP_ENT_FATURA_DETAY " & lngbelgekod
        Dim detaylar As List(Of DataRow) = lb.QueryDondurPan(detayquery)

        Dim k As Integer = 0
        For Each detay In detaylar
            k = k + 1
            If mikro.MikroUrunKontrol(detay("TXTTCPURUNKOD")) = False Then
                hatavar = True
                lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, detay("TXTTCPURUNKOD") & " Ürün Kodu Bulunamadı->" & detay("TXTTCPURUNKOD"), 0)
            End If
            If mikro.MikroBirimKontrol(detay("TXTTCPBIRIMACIKLAMA")) = False Then
                hatavar = True
                lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, detay("BYTTCPBIRIMSIRA") & "-> " & detay("TXTTCPBIRIMACIKLAMA") & " Birim Kodu Bulunamadı->" & detay("TXTTCPBIRIMACIKLAMA"), 0)
            End If
        Next


        If hatavar = False Then
            carihareketid = mikro.CariHesapBaslikKaydet("FATURA", firmano, subeno, evrakseri, evraksira + 1, belgetarih, belgeno, stkodu, carikodu, bruttutar, nettutar, kdvtutar1, kdvtutar2, kdvtutar3, kdvtutar4, iskontotutar, ebelge, kaynakbelgeno, aciklama)
            If carihareketid.ToString = "" Then
                lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Fatura Başlık bilgileri kaydedilmesi sırasında hata  INSERT SORUNU !!!.", 0)
                hatavar = True
            Else
                lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Fatura Başlık bilgileri kaydedildi.", 1)
                Pan.PanoramaFaturaAktarildiYap(kaynakbelgeno)
            End If

            Dim stoksiralist As List(Of DataRow) = lb.QueryDondurMikro("SELECT MAX(sth_evrakno_sira) as stoksirano FROM STOK_HAREKETLERI WHERE sth_evraktip=4 and sth_subeno=" & subeno & " and sth_firmano=" & firmano)
            Dim stoksira As Integer = 0
            If IsDBNull(stoksiralist(0)("stoksirano")) = False Then
                stoksira = stoksiralist(0)("stoksirano")
            End If

            Dim i As Integer = 0
            For Each detay In detaylar
                i = i + 1
                'Birim Sırayı Öğren
                'mikro.SatisFaturaDetayKaydet(firmano, subeno, carihareketid, belgetarih, belgeno, stoksira + i, detay("LNGKALEMSIRA"), detay("TXTTCPURUNKOD"), carikodu, stkodu, detay("DBLMIKTAR"), detay("BYTTCPBIRIMSIRA"), detay("DBLTUTAR"), detay("DBLKDVTUTAR"), detay("LNGDEPOKOD"), detay("LNGDEPOKOD"))
                mikro.SatisFaturaDetayKaydet(firmano, subeno, carihareketid, belgetarih, belgeno, stoksira + i, detay("LNGKALEMSIRA"), detay("TXTTCPURUNKOD"), carikodu, stkodu, detay("DBLMIKTAR"), detay("BYTTCPBIRIMSIRA"), detay("DBLTUTAR"), detay("DBLKDVTUTAR"), detay("DBLKDVORANISIRA"), detay("DBLISKONTOTUTARI"), detay("LNGDEPOKOD"), detay("LNGDEPOKOD"), detay("ISKTUTAR1"), detay("ISKTUTAR2"), detay("ISKTUTAR3"), detay("ISKTUTAR4"), detay("ISKTUTAR5"), detay("ISKTUTAR6"))
            Next
            lb.Entlogyaz(kaynakbelgeno, hedefbelgeno, "Fatura Detay bilgileri kaydedildi.", 1)
        End If

    End Function




End Class