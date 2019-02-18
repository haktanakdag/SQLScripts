Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.IO
Imports System.Configuration
Imports System.Data.OleDb
Imports System.IO.Compression
Imports System.Xml
Imports System.Net
Imports System.Web
Imports System.Net.Mail
Imports System.Globalization
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports System.Security.Cryptography

Public Class voilib
    Public Function ParasalDegerDondur(ByVal deger As String) As String
        Return String.Format("{0:n}", CType(deger, Double))
    End Function

    Function tarihDondurKerzz(ByVal tarih As String) As String
        Dim yil As String = ""
        Dim ay As String = ""
        Dim gun As String = ""
        yil = Convert.ToDateTime(tarih).Year.ToString
        ay = Convert.ToDateTime(tarih).Month.ToString
        If (ay.Count = 1) Then
            ay = "0" & ay
        End If
        gun = Convert.ToDateTime(tarih).Day.ToString
        If (gun.Count = 1) Then
            gun = "0" & gun
        End If

        Return yil & "-" & ay & "-" & gun
    End Function
    Function tarihDondurPan(ByVal tarih As String) As String
        Dim yil As String = ""
        Dim ay As String = ""
        Dim gun As String = ""
        yil = Convert.ToDateTime(tarih).Year.ToString
        ay = Convert.ToDateTime(tarih).Month.ToString
        If (ay.Count = 1) Then
            ay = "0" & ay
        End If
        gun = Convert.ToDateTime(tarih).Day.ToString
        If (gun.Count = 1) Then
            gun = "0" & gun
        End If

        Return yil & "-" & ay & "-" & gun
    End Function

    Function tarihDondurNetsis(ByVal tarih As String) As String
        Dim yil As String = ""
        Dim ay As String = ""
        Dim gun As String = ""
        yil = Convert.ToDateTime(tarih).Year.ToString
        ay = Convert.ToDateTime(tarih).Month.ToString
        If (ay.Count = 1) Then
            ay = "0" & ay
        End If
        gun = Convert.ToDateTime(tarih).Day.ToString
        If (gun.Count = 1) Then
            gun = "0" & gun
        End If

        Return gun & "/" & ay & "/" & yil
    End Function
    Public Function tarihYYYYMMDDondur(txttarih As String) As String
        Dim tarih As String = ""
        If txttarih = "" Then
            tarih = "00.00.0000"
        Else
            Dim yil As String = Convert.ToDateTime(txttarih).Year
            Dim ay As String = Convert.ToDateTime(txttarih).Month
            If ay.Count = 1 Then
                ay = "0" & ay
            End If
            Dim gun As String = Convert.ToDateTime(txttarih).Day
            If gun.Count = 1 Then
                gun = "0" & gun
            End If
            tarih = yil & ay & gun
        End If
        Return tarih
    End Function
    Public Function tarihDDMMYYYYDondur(txttarih As String) As String
        Dim tarih As String = ""
        If txttarih.ToString = "" Then
            tarih = "00.00.0000"
        Else
            Dim yil As String = Convert.ToDateTime(txttarih).Year
            Dim ay As String = Convert.ToDateTime(txttarih).Month
            If ay.Count = 1 Then
                ay = "0" & ay
            End If
            Dim gun As String = Convert.ToDateTime(txttarih).Day
            If gun.Count = 1 Then
                gun = "0" & gun
            End If
            tarih = gun & "." & ay & "." & yil
        End If
        Return tarih
    End Function
    Public Function FloadKontrol(deger As String) As String
        deger = deger.Replace(" ", "")
        If deger.IndexOf(",") > 0 Then
            deger = deger.Replace(",", ".")
        End If
        Dim decVal As Decimal

        If (Decimal.TryParse(deger, decVal)) = False Then
            deger = ""
            Return deger
        Else
            If deger.IndexOf(".") > 0 Then
                Return deger
            Else
                Return deger & ".00"
            End If
        End If
    End Function
    Public Function kolonadiduzelt(ByVal kelimecik As String) As String

        kelimecik = kelimecik.Replace("ö", "o")
        kelimecik = kelimecik.Replace("ü", "u")
        kelimecik = kelimecik.Replace("ğ", "g")
        kelimecik = kelimecik.Replace("ş", "s")
        kelimecik = kelimecik.Replace("ı", "i")
        kelimecik = kelimecik.Replace("ç", "c")
        kelimecik = kelimecik.Replace("Ö", "O")
        kelimecik = kelimecik.Replace("Ü", "U")
        kelimecik = kelimecik.Replace("Ğ", "G")
        kelimecik = kelimecik.Replace("Ş", "S")
        kelimecik = kelimecik.Replace("İ", "I")
        kelimecik = kelimecik.Replace("Ç", "C")
        kelimecik = kelimecik.Replace(" ", "_")
        kelimecik = kelimecik.Replace("'", " ")
        Return kelimecik
    End Function
    Public Function stringduzelt(ByVal kelimecik As String) As String
        kelimecik = kelimecik.Replace("'", " ")
        Return kelimecik
    End Function
    
    Public Shared Function GetDateTime() As DateTime
        Dim dateTime As DateTime = DateTime.MinValue
        Dim request As System.Net.HttpWebRequest = DirectCast(System.Net.WebRequest.Create("http://www.microsoft.com"), System.Net.HttpWebRequest)
        request.Method = "GET"
        request.Accept = "text/html, application/xhtml+xml, */*"
        request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)"
        request.ContentType = "application/x-www-form-urlencoded"
        request.CachePolicy = New System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore)
        Dim response As System.Net.HttpWebResponse = DirectCast(request.GetResponse(), System.Net.HttpWebResponse)
        If response.StatusCode = System.Net.HttpStatusCode.OK Then
            Dim todaysDates As String = response.Headers("date")
            dateTime = DateTime.ParseExact(todaysDates, "ddd, dd MMM yyyy HH:mm:ss 'GMT'", System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat, System.Globalization.DateTimeStyles.AssumeUniversal)
        End If
        Return dateTime
    End Function
    Public Function GetDateTimeSTR() As String
        Dim dateTime As DateTime = DateTime.MinValue
        Dim request As System.Net.HttpWebRequest = DirectCast(System.Net.WebRequest.Create("http://www.microsoft.com"), System.Net.HttpWebRequest)
        request.Method = "GET"
        request.Accept = "text/html, application/xhtml+xml, */*"
        request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)"
        request.ContentType = "application/x-www-form-urlencoded"
        request.CachePolicy = New System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore)
        Dim response As System.Net.HttpWebResponse = DirectCast(request.GetResponse(), System.Net.HttpWebResponse)
        If response.StatusCode = System.Net.HttpStatusCode.OK Then
            Dim todaysDates As String = response.Headers("date")
            dateTime = DateTime.ParseExact(todaysDates, "ddd, dd MMM yyyy HH:mm:ss 'GMT'", System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat, System.Globalization.DateTimeStyles.AssumeUniversal)
        End If
        Dim tarihsaatstr As String = dateTime
        tarihsaatstr = tarihsaatstr.Replace(".", "")
        tarihsaatstr = tarihsaatstr.Replace(":", "")
        tarihsaatstr = tarihsaatstr.Replace(" ", "")
        Return tarihsaatstr
    End Function
    
    Public Shared Function QueryDondur(ByVal q As String) As List(Of DataRow)
        Dim Dbl As New DBLib
        Dim Connection As New SqlConnection(Dbl.gGetConnectionStringEntegrator())
        Connection.Open()
        Dim cmd As New SqlCommand
        cmd.Connection = Connection
        cmd.CommandType = CommandType.Text
        cmd.CommandText = q

        Dim sda As New SqlDataAdapter()
        sda.SelectCommand = cmd
        Dim dt As New DataTable()
        sda.Fill(dt)
        Dim dr As New List(Of DataRow)()

        For Each row As DataRow In dt.Rows
            dr.Add(CType(row, DataRow))
        Next row

        Connection.Close()
        Return dr
    End Function
    Public Shared Function QueryDondurMikro(ByVal q As String) As List(Of DataRow)
        Dim Dbl As New DBLib
        Dim Connection As New SqlConnection(Dbl.gGetConnectionStringMikro())
        Connection.Open()
        Dim cmd As New SqlCommand
        cmd.Connection = Connection
        cmd.CommandType = CommandType.Text
        cmd.CommandText = q

        Dim sda As New SqlDataAdapter()
        sda.SelectCommand = cmd
        Dim dt As New DataTable()
        sda.Fill(dt)
        Dim dr As New List(Of DataRow)()

        For Each row As DataRow In dt.Rows
            dr.Add(CType(row, DataRow))
        Next row

        Connection.Close()
        Return dr
    End Function

    Public Shared Function QueryDondurMikroAyar(ByVal q As String) As List(Of DataRow)
        Dim Dbl As New DBLib
        Dim Connection As New SqlConnection(Dbl.gGetConnectionStringMikroAyar())
        Connection.Open()
        Dim cmd As New SqlCommand
        cmd.Connection = Connection
        cmd.CommandType = CommandType.Text
        cmd.CommandText = q

        Dim sda As New SqlDataAdapter()
        sda.SelectCommand = cmd
        Dim dt As New DataTable()
        sda.Fill(dt)
        Dim dr As New List(Of DataRow)()

        For Each row As DataRow In dt.Rows
            dr.Add(CType(row, DataRow))
        Next row

        Connection.Close()
        Return dr
    End Function
    Public Shared Function QueryDondurPan(ByVal q As String) As List(Of DataRow)
        Dim Dbl As New DBLib
        Dim Connection As New SqlConnection(Dbl.gGetConnectionStringPan())
        Connection.Open()
        Dim cmd As New SqlCommand
        cmd.Connection = Connection
        cmd.CommandType = CommandType.Text
        cmd.CommandText = q

        Dim sda As New SqlDataAdapter()
        sda.SelectCommand = cmd
        Dim dt As New DataTable()
        sda.Fill(dt)
        Dim dr As New List(Of DataRow)()

        For Each row As DataRow In dt.Rows
            dr.Add(CType(row, DataRow))
        Next row

        Connection.Close()
        Return dr
    End Function
    Public Sub GridFormatting(ByVal grdKontrol As DataGridView)
        'grdRptTanim.ForeColor = Color.Black
        'grdRptTanim.BackgroundColor = Color.AliceBlue
        'grdRptTanim.AlternatingRowsDefaultCellStyle.BackColor = Color.Brown
        'grdRptTanim.AlternatingRowsDefaultCellStyle.ForeColor = Color.DodgerBlue
        'grdRptTanim.ColumnHeadersDefaultCellStyle.ForeColor = Color.CadetBlue
        'grdRptTanim.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGoldenrod
        grdKontrol.EnableHeadersVisualStyles = False
        grdKontrol.AutoGenerateColumns = True
        grdKontrol.AllowUserToAddRows = False
        grdKontrol.ReadOnly = True
        grdKontrol.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grdKontrol.Refresh()
    End Sub

    Public Function IsValidEmail(ByVal email As String) As Boolean
        Try
            Dim mail = New System.Net.Mail.MailAddress(email)
            Return True
        Catch
            Return False
        End Try
    End Function

    Public Function Entlogyaz(ByVal kaynakbelgekod As String, ByVal hedefbelgekod As String, ByVal aciklama As String, ByVal aktarimdurum As Int16)
        Dim cmd As New SqlCommand
        Dim textaciklama As String = aciklama.Replace("'", "")
        Dim Sql = "INSERT INTO VOID_ENT_LOG(KAYNAKBELGEKOD,HEDEFBELGEKOD, ACIKLAMA, TARIH,AKTARIMDURUM) values ('" & kaynakbelgekod & "','" & hedefbelgekod & "','" & textaciklama & "', GETDATE() , " & aktarimdurum & " )"
        Dim veri As New DBLib
        Using connection As New SqlConnection(veri.gGetConnectionStringEntegrator())
            Try
                connection.Open()
                cmd.Connection = connection
                cmd.CommandText = Sql
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show("SQL Message" & ex.Message, "ERROR")
            Finally
                connection.Close()
            End Try
        End Using
    End Function
End Class
