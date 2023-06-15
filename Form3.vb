
Imports QRCoder
Imports System.Text
Imports System.Drawing.Printing
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO
Imports System.Configuration
Imports System.Collections.Specialized





Public Class Form3

    Private printFont As Font
    Private streamToPrint As StreamReader
    Public sAttr As String
    Public sfichinv As String
    Public sfichcsv As String
    Public Sfichlogo As String
    Public sadobe As String
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        '   sAttr = ConfigurationManager.AppSettings.Get("Key0") 'voir app.config
        sfichinv = ConfigurationManager.AppSettings.Get("Key3") 'voir app.config
        sfichcsv = ConfigurationManager.AppSettings.Get("Key4") 'voir app.config
        Sfichlogo = ConfigurationManager.AppSettings.Get("Key5") 'voir app.config
        sadobe = ConfigurationManager.AppSettings.Get("Key4") 'voir app.config
        Dim produit As String
        produit = "Moteurs Sirem"
        Using MyReader As New Microsoft.VisualBasic.
                      FileIO.TextFieldParser(
                        sfichinv)


            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(";")
            Dim currentRow As String()
            Dim textFileStream As New FileStream(Environment.CurrentDirectory + "\csv\fichier.csv", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
            Dim ecriture As New StreamWriter(textFileStream)
            Dim liste As New List(Of String)
            Dim pgsiz1 As New iTextSharp.text.Rectangle(PageSize.A4)
            Dim pdfDoc1 As New Document(pgsiz1)
            Dim pdfWrite1 As PdfWriter = PdfAWriter.GetInstance(pdfDoc1, New FileStream(Environment.CurrentDirectory + "\pdf\fichier.pdf", FileMode.Create))
            Dim y As Integer
            y = 0
            While Not MyReader.EndOfData
                Try

                    currentRow = MyReader.ReadFields()

                    Dim currentField As String
                    For Each currentField In currentRow
                        If Not (liste.Contains(currentField)) Then
                            Dim currentField1 = "=" & """" & currentField & """"
                            currentField1 = Replace(currentField1, " ", "")
                            currentField1 = Replace(currentField1, "/", "_")
                            liste.Add(currentField1)
                            pdfDoc1.Open()
                            If currentField = "1" Then
                                currentField = currentField + 1
                            Else
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                currentField = Replace(currentField, " ", "")
                                currentField = Replace(currentField, "/", "_")
                                'Nombre de caracteres du numéro de serie 'donnée,n° caractere de debut ,nb de caractere a garder + suppression des espaces

                                ' currentField = Mid(currentField, 15, 10)
                                currentField = Replace(currentField, " ", "")
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                Dim gen As New QRCodeGenerator
                                Dim data = gen.CreateQrCode(currentField, QRCodeGenerator.ECCLevel.Q)

                                Dim code As New qrcode(data)
                                Dim bitMap As Bitmap = code.GetGraphic(4)
                                PictureBox1.Image = code.GetGraphic(6)
                                PictureBox1.Image = bitMap
                                PictureBox1.Height = bitMap.Height
                                PictureBox1.Width = bitMap.Width

                                PictureBox1.Image.Save(Sfichlogo)
                                Threading.Thread.Sleep(1000)

                                'Create a new PDF document
                                Dim pgsiz As New iTextSharp.text.Rectangle(142, 71)

                                '  Dim pdfDoc As New Document(pgsiz, 0, 0, 0, 0)  ' 10.0F
                                Dim pdfDoc As New Document(pgsiz)

                                ' Dim cb As iTextSharp.text.pdf.PdfContentByte = Nothing

                                Dim pdfWrite As PdfWriter = PdfAWriter.GetInstance(pdfDoc, New FileStream(Environment.CurrentDirectory + "\pdf\ " & currentField & ".pdf", FileMode.Create))

                                pdfDoc.Open()
                                pdfDoc.NewPage()

                                Dim table3 As PdfPTable = New PdfPTable(2)

                                Dim cell2 As PdfPCell = New PdfPCell()

                                Dim Image = iTextSharp.text.Image.GetInstance(Sfichlogo)
                                Image.SetAbsolutePosition(50, 20)
                                Image.ScaleToFit(50, 50)


                                Dim cb As PdfContentByte = pdfWrite.DirectContent

                                ' select the font properties
                                Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
                                cb.SetColorFill(BaseColor.DARK_GRAY)
                                cb.SetFontAndSize(bf, 8)

                                ' write the text in the pdf content
                                cb.BeginText()
                                Dim text As String = currentField
                                ' put the alignment and coordinates here
                                cb.ShowTextAligned(1, text, 75, 10, 0)
                                cb.EndText()

                                pdfDoc.Add(Image)
                                pdfDoc.NewPage()





                                Image.SetAbsolutePosition(50, 20)
                                Image.ScaleToFit(50, 50)

                                cb.SetColorFill(BaseColor.DARK_GRAY)
                                cb.SetFontAndSize(bf, 8)

                                ' write the text in the pdf content
                                cb.BeginText()

                                ' put the alignment and coordinates here
                                cb.ShowTextAligned(1, text, 75, 10, 0)
                                cb.EndText()
                                pdfDoc.Add(Image)

                                pdfDoc.Close()


                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ''''impression sur Zebra
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                ' impzebra(currentField)

                                '  Dim pathToExecutable As String = sadobe
                                ' Dim sReport = "\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\pdf \ " & currentField & ".pdf" 'Complete name/path of PDF file    
                                ' sAttr = ConfigurationManager.AppSettings.Get("Key1")
                                '  Dim SPrinter = sAttr  'Name Of printer    
                                '  Dim starter As New ProcessStartInfo(pathToExecutable, "/t " + sReport + " " + SPrinter + "")
                                'Dim starter As New ProcessStartInfo(pathToExecutable + sReport + " " + SPrinter + "")
                                '  Dim Process As New Process()
                                '   Process.StartInfo.CreateNoWindow = True

                                '   Process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                '   Process.StartInfo = starter
                                '    Process.Start()

                                '  Process.WaitForExit(10)

                                '  Process.Dispose()
                                ' impression1("\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\pdf\" & currentField & ".pdf")
                                Dim stream = New StreamReader(Environment.CurrentDirectory + "\pdf\ " & currentField & ".pdf")





                                Dim imprimante As New PrintDocument ()

                                imprimante.PrinterSettings.PrinterName = "HP LaserJet Pro MFP M426-M427 PCL 6 (Copie 1)"
                                imprimante.Print()

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                            End If

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

                        End If



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
            'impression liste
            impression("\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\pdf\fichier.pdf")
            '  imp1()

            '''''''''''''''''''''''''''''''''''''''''''''''''''
        End Using

        deplacement(produit)
        Threading.Thread.Sleep(10000)
        MsgBox("appuyer sur une touche une fois toutes les étiquettes sorties")
        Kill(Environment.CurrentDirectory + "\pdf\*.pdf")
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim printers As System.Drawing.Printing.PrinterSettings.StringCollection
        printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters()
        For x As Integer = 0 To printers.Count - 1

        Next
    End Sub
End Class