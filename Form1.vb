
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO
Imports System.ComponentModel

Imports System.Configuration


Public Class Form1

    Public sAttr As String
    Public sfichinv As String

    Public Sfichlogo As String


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        impression("C:\pdf\fichier.pdf")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' sAttr = ConfigurationManager.AppSettings.Get("Key1") 'voir app.config
        sfichinv = ConfigurationManager.AppSettings.Get("Key3") 'voir app.config
        '  sfichcsv = ConfigurationManager.AppSettings.Get("Key4") 'voir app.config
        '   Sfichlogo = ConfigurationManager.AppSettings.Get("Key5") 'voir app.config
        Dim produit As String
        produit = "Coffrets sielmel"
        Using MyReader As New Microsoft.VisualBasic.
                      FileIO.TextFieldParser(
                        sfichinv)


            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(";")
            Dim currentRow As String()
            Dim textFileStream As New FileStream("\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\csv\fichier.csv", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
            Dim ecriture As New StreamWriter(textFileStream)
            Dim liste As New List(Of String)
            Dim pgsiz1 As New iTextSharp.text.Rectangle(PageSize.A4)
            Dim pdfDoc1 As New Document(pgsiz1)
            Dim pdfWrite1 As PdfWriter = PdfAWriter.GetInstance(pdfDoc1, New FileStream("\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\pdf\fichier.pdf", FileMode.Create))
            Dim y As Integer
            y = 0
            While Not MyReader.EndOfData
                Try

                    currentRow = MyReader.ReadFields()

                    Dim currentField As String
                    For Each currentField In currentRow
                        '  If Not (liste.Contains(currentField)) Then

                        Dim currentField1 = "=" & """" & currentField & """"
                        ' currentField1 = Replace(currentField1, " ", "")
                        currentField1 = Replace(currentField1, "/", "_")
                        liste.Add(currentField1)

                        pdfDoc1.Open()
                        If currentField = "1" Then
                            currentField = currentField + 1
                        Else
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                            'Nombre de caracteres du numéro de serie 'donnée,n° caractere de debut ,nb de caractere a garder + suppression des espaces

                            ' currentField = Mid(currentField, 2, 2)
                            'currentField = Replace(currentField, " ", "")

                            currentField = Replace(currentField, "/", "_")
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        End If

                        ' Dim Image = iTextSharp.text.Image.GetInstance(Sfichlogo)
                        ' Image.SetAbsolutePosition(50, 20)
                        '  Image.ScaleToFit(40, 40)
                        y = y + 1
                        '  liste.Add(currentField)
                        pdfDoc1.Add(New Phrase(currentField))
                        ' pdfDoc1.Add(Image)
                        If y = 40 Then
                            pdfDoc1.NewPage()
                        End If
                        If y = 80 Then
                            pdfDoc1.NewPage()
                        End If
                        If y = 120 Then
                            pdfDoc1.NewPage()
                        End If
                        If y = 160 Then
                            pdfDoc1.NewPage()
                        End If
                        '  End If



                    Next

                Catch ex As Microsoft.VisualBasic.
                    FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
            "is not valid and will be skipped.")
                End Try

            End While
            pdfDoc1.Close()
            Dim sortQuery = From word In liste
                            Order By Val(word) Ascending
                            Select word
            For Each elements As String In sortQuery

                ecriture.WriteLine(elements)
            Next


            ecriture.Close()
            '''''''''''''''''''''''''''''''''''''''''''''''''''

            impression("\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\pdf\fichier.pdf")

            '''''''''''''''''''''''''''''''''''''''''''''''''''
        End Using
        ' deplacement(produit)
    End Sub
End Class
