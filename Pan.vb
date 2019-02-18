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
Public Class Pan

    Function PanoramaFaturaAktarildiYap(faturakod As String)
        Dim cmd As New SqlCommand
        Dim Sql = "UPDATE TBLMSDFATURA SET BYTAKTARILDI=1 WHERE LNGBELGEKOD =" & faturakod
        Dim veri As New DBLib
        Using connection As New SqlConnection(veri.gGetConnectionStringPan())
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
