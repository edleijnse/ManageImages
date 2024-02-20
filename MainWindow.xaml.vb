Imports System.IO
Imports System.Text
Imports IniParser
Imports IniParser.Model

Public Class MainWindow
    Private ReadOnly parser As New FileIniDataParser()

    Public Sub New()
        InitializeComponent()
        AddHandler Me.Loaded, AddressOf Window_Loaded
        LoadButtonContentsFromIniFile()
    End Sub

    Public Sub LoadButtonContentsFromIniFile()
        Dim filePath As String = GetIniFilePath()
        Try
            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)
                Dim buttonSection = data("Buttons")
                AssignContentToButtons(buttonSection)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
                ' Create file
                ' Create a new file     
                Using sw As StreamWriter = File.CreateText(filePath)
                    ' Add some text to file
                    sw.WriteLine("[Buttons]")
                    sw.WriteLine("Button1 = DROP application here")
                    sw.WriteLine("Button2 = DROP application here")
                    sw.WriteLine("Button3 = DROP application here")
                    sw.WriteLine("Button4 = DROP application here")
                End Using
                Dim data = parser.ReadFile(filePath)
                Dim buttonSection = data("Buttons")
                AssignContentToButtons(buttonSection)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Function GetButton1Content() As String
        Return Button1.Content.ToString()
    End Function
    Public Function GetButton2Content() As String
        Return Button2.Content.ToString()
    End Function
    Public Function GetButton3Content() As String
        Return Button3.Content.ToString()
    End Function
    Public Function GetButton4Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton5Content() As String
        Return Button1.Content.ToString()
    End Function
    Public Function GetButton6Content() As String
        Return Button2.Content.ToString()
    End Function
    Public Function GetButton7Content() As String
        Return Button3.Content.ToString()
    End Function
    Public Function GetButton8Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton9Content() As String
        Return Button1.Content.ToString()
    End Function
    Public Function GetButton10Content() As String
        Return Button2.Content.ToString()
    End Function
    Public Function GetButton11Content() As String
        Return Button3.Content.ToString()
    End Function
    Public Function GetButton12Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton13Content() As String
        Return Button1.Content.ToString()
    End Function
    Public Function GetButton14Content() As String
        Return Button2.Content.ToString()
    End Function
    Public Function GetButton15Content() As String
        Return Button3.Content.ToString()
    End Function
    Public Function GetButton16Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton17Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton18Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton19Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Function GetButton20Content() As String
        Return Button4.Content.ToString()
    End Function
    Public Sub AssignContentToButtons(buttonSection As KeyDataCollection)
        Button1.Content = Path.GetFileName(buttonSection("Button1"))
        Button2.Content = Path.GetFileName(buttonSection("Button2"))
        Button3.Content = Path.GetFileName(buttonSection("Button3"))
        Button4.Content = Path.GetFileName(buttonSection("Button4"))
        Button5.Content = Path.GetFileName(buttonSection("Button5"))
        Button6.Content = Path.GetFileName(buttonSection("Button6"))
        Button7.Content = Path.GetFileName(buttonSection("Button7"))
        Button8.Content = Path.GetFileName(buttonSection("Button8"))
        Button9.Content = Path.GetFileName(buttonSection("Button9"))
        Button10.Content = Path.GetFileName(buttonSection("Button10"))
        Button11.Content = Path.GetFileName(buttonSection("Button11"))
        Button12.Content = Path.GetFileName(buttonSection("Button12"))
        Button13.Content = Path.GetFileName(buttonSection("Button13"))
        Button14.Content = Path.GetFileName(buttonSection("Button14"))
        Button15.Content = Path.GetFileName(buttonSection("Button15"))
        Button16.Content = Path.GetFileName(buttonSection("Button16"))
        Button17.Content = Path.GetFileName(buttonSection("Button17"))
        Button18.Content = Path.GetFileName(buttonSection("Button18"))
        Button19.Content = Path.GetFileName(buttonSection("Button19"))
        Button20.Content = Path.GetFileName(buttonSection("Button20"))
    End Sub

    Public Function GetIniFilePath() As String
        Dim rootPath As String = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName
        Return Path.Combine(rootPath, "startapplications.ini")
    End Function

    Public Sub ChangeButtonContent(myButton As String, buttonContent As String)
        Dim filePath As String = GetIniFilePath()
        Try
            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)
                data("Buttons")(myButton) = buttonContent
                parser.WriteFile(filePath, data)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Me.Top = SystemParameters.WorkArea.Height - Me.Height
        Me.Left = (SystemParameters.WorkArea.Width - Me.Width) / 2
    End Sub

    Private Sub Window_KeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.Key
            Case Key.Left
                Me.Left -= 200
            Case Key.Right
                Me.Left += 200
            Case Key.Up
                Me.Top -= 200
            Case Key.Down
                Me.Top += 200
        End Select
    End Sub

    Private Sub BtnDropFile_Drop(sender As Object, e As DragEventArgs)
        Dim button = CType(sender, Button)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files IsNot Nothing AndAlso files.Length <> 0 Then
                Dim lastPart As String = Path.GetFileName(files(0))
                button.Content = lastPart
                ChangeButtonContent(button.Name, files(0))
            End If
        End If
    End Sub

    Private Sub ButtonClickHandler(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim myApplicationHandler As New ApplicationHandler()
        Select Case button.Content
            Case "Convert Raw"
                myApplicationHandler.ConvertRaw()
            Case "Start Excel"
                myApplicationHandler.startExcel()
            Case "Start OneNote"
                myApplicationHandler.StartOneNote()
            Case "Start Outlook"
                myApplicationHandler.StartOutlook()
            Case "magick"
                myApplicationHandler.StartMagick()
            Case Else
                StartCustomApplication(button.Name, myApplicationHandler)
        End Select
    End Sub

    Private Sub StartCustomApplication(buttonName As String, applicationHandler As ApplicationHandler)
        Dim filePath As String = GetIniFilePath()
        Try
            If File.Exists(filePath) Then
                Dim data = parser.ReadFile(filePath)
                Dim myApplication = data("Buttons")(buttonName)
                applicationHandler.StartApplication(myApplication)
            Else
                MessageBox.Show($"File: {filePath} does not exist.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub Button_Click(sender As Object, e As RoutedEventArgs)
        ' Your event handling logic goes here.
    End Sub
End Class