
Imports System.IO
Imports System.Configuration
Imports System.Drawing.Printing

Module Module1


    'test
    Public sAttr As String
    Public sAttr1 As String
    Public isEmpty As Boolean = True



    Sub impression(ByVal pathFile As String)
        sAttr = ConfigurationManager.AppSettings.Get("Key1")
        Dim pathToExecutable As String = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
        Dim sReport = pathFile   'Complete name/path of PDF file    
        Dim SPrinter = sAttr  'Name Of printer 

        Dim starter As New ProcessStartInfo(pathToExecutable, "/t " + sReport + " " + SPrinter + "")
        Dim Process As New Process()
        Process.StartInfo = starter
        Process.Start()
        'try and close the process with 20 seconds delay
        System.Threading.Thread.Sleep(3)
        ' Process.CloseMainWindow()
        Dim iLoop As Int16 = 0
        'check the process has exited or not
        If Process.HasExited = False Then
            'if not then loop for 100 time to try and close the process'with 10 seconds delay
            While Not Process.HasExited
                System.Threading.Thread.Sleep(1)
                Process.CloseMainWindow()
                iLoop = CShort(iLoop + 1)
                If iLoop >= 100 Then
                    Exit While
                End If
            End While
        End If
        Process.Close()
        Process.Dispose()
        Process = Nothing
        starter = Nothing
    End Sub
    Sub impression1(ByVal pathFile As String)
        sAttr1 = ConfigurationManager.AppSettings.Get("Key0")
        Dim pathToExecutable As String = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
        Dim sReport = pathFile   'Complete name/path of PDF file    
        Dim SPrinter = sAttr1  'Name Of printer 

        Dim starter As New ProcessStartInfo(pathToExecutable, "/t " + sReport + " " + "ZDesigner ZM400 200 dpi (ZPL) (Copie 1)" + "")
        Dim Process As New Process()
        Process.StartInfo = starter
        Process.Start()
        'try and close the process with 20 seconds delay
        System.Threading.Thread.Sleep(3)
        ' Process.CloseMainWindow()
        Dim iLoop As Int16 = 0
        'check the process has exited or not
        If Process.HasExited = False Then
            'if not then loop for 100 time to try and close the process'with 10 seconds delay
            While Not Process.HasExited
                System.Threading.Thread.Sleep(1)
                Process.CloseMainWindow()
                iLoop = CShort(iLoop + 1)
                If iLoop >= 100 Then
                    Exit While
                End If
            End While
        End If
        Process.Close()
        Process.Dispose()
        Process = Nothing
        starter = Nothing
    End Sub
    Sub deplacement(produit As String)

        Dim dte As String
        dte = Now
        dte = Replace(dte, " ", "_")
        dte = Replace(dte, "/", "_")
        dte = Replace(dte, ":", "_")

        My.Computer.FileSystem.MoveFile(Environment.CurrentDirectory + "\csv\fichier.csv",
    "\\SRVDC3.next-pool.local\abriblue\PRODUCTION\reception\" & produit & "_" & dte & "_fichier.csv",
    FileIO.UIOption.AllDialogs,
    FileIO.UICancelOption.ThrowException)
    End Sub





End Module
