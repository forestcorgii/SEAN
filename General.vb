Imports System.Windows.Forms
Imports System.IO
Imports System.Text


Public Class General
    Private Const myComputerName As String = "IDCFC-0192"
    Public Shared Sub DebugMessage(msg As String)
        If My.Computer.Name = myComputerName Then
            MsgBox(msg)
        End If
    End Sub

    Public Class Math
        Public Shared Function Between(val As Integer, min As Integer, max As Integer) As Boolean
            Return val >= min And val <= max
        End Function
    End Class



    Public Class Dialogs
        Public Shared Function OpenFileDialogShortcut(Optional filter As String = "*", Optional title As String = "", Optional multiSelect As Boolean = True, Optional filename As String = "")
            Dim tmpOFD As New OpenFileDialog
            With tmpOFD
                .Title = title
                .Multiselect = multiSelect
                .FileName = filename
                .Filter = filter
                If .ShowDialog() = DialogResult.OK Then
                    If .FileNames.Count = 1 Then
                        Return ConvertToUNCPath(.FileName)
                    Else : Return (From res As String In .FileNames Select ConvertToUNCPath(res)).ToArray
                    End If
                End If
            End With
            Return Nothing
        End Function

        Public Shared Function FolderBrowserDialogShortcut(Optional description As String = "") As String
            Dim tmpFBD As New FolderBrowserDialog
            With tmpFBD
                .Description = description
                If .ShowDialog() = DialogResult.OK Then
                    Return ConvertToUNCPath(.SelectedPath)
                End If
            End With
            Return Nothing
        End Function
    End Class
    Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, _
           ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer
    ''' <summary>
    ''' THIS CONVERTS DRIVE LETTER TO UNCPATH
    ''' </summary>
    ''' <param name="sFilePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertToUNCPath(ByVal sFilePath As String) As String
        If sFilePath.Trim = "" Or Not sFilePath.Length >= 3 Then Return ""
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo
        Dim DriveType, Ctr As Integer
        Dim DriveLtr, UNCName As String
        Dim StrBldr As New StringBuilder
        If sFilePath.StartsWith("\\") Then Return sFilePath
        UNCName = Space(160)
        ConvertToUNCPath = ""
        DriveLtr = sFilePath.Substring(0, 3)
        For Each d In allDrives
            If d.Name = DriveLtr Then
                DriveType = d.DriveType
                Exit For
            End If
        Next
        If DriveType = 4 Then
            Ctr = WNetGetConnection(sFilePath.Substring(0, 2), UNCName, UNCName.Length)
            If Ctr = 0 Then
                UNCName = UNCName.Trim
                For Ctr = 0 To UNCName.Length - 1
                    Dim SingleChar As Char = UNCName(Ctr)
                    Dim asciiValue As Integer = Asc(SingleChar)
                    If asciiValue > 0 Then
                        StrBldr.Append(SingleChar)
                    Else
                        Exit For
                    End If
                Next
                StrBldr.Append(sFilePath.Substring(2))
                ConvertToUNCPath = StrBldr.ToString
            Else
                MsgBox("Cannot Retrieve UNC path" & vbCrLf & "Must Use Mapped Drive of SQLServer", MsgBoxStyle.Critical)
            End If
        Else
            Return sFilePath
        End If
    End Function
    'Public Shared Function ConvertPath(filepath As String) As String


    '    Return ""
    'End Function

End Class
