Imports System.ComponentModel


Public Class MenuButton
    Private _iconSize As Integer = 32

    Public Property IconSize As Integer
        Get
            Return _iconSize
        End Get
        Set(value As Integer)
            _iconSize = value

        End Set
    End Property

    Public Property LabelText As String
        Get
            Return Label1.Text
        End Get
        Set(value As String)
            Label1.Text = value
        End Set
    End Property

    Public Property Image As System.Drawing.Image
        Get
            Return PictureBox1.Image
        End Get
        Set(value As System.Drawing.Image)
            PictureBox1.Image = value
        End Set
    End Property

    Public Sub New(key As String)
        InitializeComponent()
        Name = key
    End Sub

    Public Sub New()
        Try
            InitializeComponent()
            Name = "Circle"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
