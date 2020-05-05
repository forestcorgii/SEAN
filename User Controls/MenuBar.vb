Imports System.ComponentModel

Public Class MenuBar
    Inherits System.Windows.Forms.Control
    Private _Menus As New MenuCollection

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    Public Property Menus As MenuCollection
        Get
            If _Menus Is Nothing Then
                MsgBox("asda")
                _Menus = New MenuCollection
            End If
            Return _Menus
        End Get
        Set(value As MenuCollection)
            _Menus = value
            '     fixMenu()
        End Set
    End Property

    'Public ReadOnly Property MenuItemCount As Integer
    '    Get
    '        Return _Menus.Count
    '    End Get
    'End Property

    Private Sub fixMenu()
        tbpnMenuSeperator.Controls.Clear()
        tbpnMenuSeperator.ColumnStyles.Clear()

        tbpnMenuSeperator.ColumnCount = _Menus.Count
        For i As Integer = 0 To _Menus.Count - 1
            Dim _menu As MenuButton = _Menus(i)
            tbpnMenuSeperator.ColumnStyles.Add(New Windows.Forms.ColumnStyle(Windows.Forms.SizeType.Absolute, _menu.Width))
            tbpnMenuSeperator.Controls.Add(_menu, 0, i)
        Next
        Me.Invalidate()
    End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '    fixMenu()
    '    '  Me.Invalidate()
    'End Sub
End Class
