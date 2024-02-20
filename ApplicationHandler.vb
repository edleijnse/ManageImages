
Public Class ApplicationHandler
    Public Sub StartWord()
        Try
            Process.Start("C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub
    Public Sub ConvertRaw()
        Try
            Process.Start("C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub

    Public Sub StartExcel()
        Try
            Process.Start("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub
    
    Public Sub StartOneNote()
        Try
            Process.Start("C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE")
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub
    
    Public Sub StartOutlook()
        Try
            Process.Start("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE")
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub
    Public Sub StartMagick()
        Dim batFilePath As String = "C:\Users\edlei\startmagickdemo.bat"
        Try
            Process.Start(batFilePath)
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub
    Public Sub StartApplication(myApplication)
        Try
            Process.Start(myApplication)
        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub
End Class