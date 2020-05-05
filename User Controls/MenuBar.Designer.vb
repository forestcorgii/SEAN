<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MenuBar
    Inherits System.Windows.Forms.Control

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.tbpnMenuSeperator = New System.Windows.Forms.TableLayoutPanel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'tbpnMenuSeperator
        '
        Me.tbpnMenuSeperator.ColumnCount = 1
        Me.tbpnMenuSeperator.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tbpnMenuSeperator.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbpnMenuSeperator.Location = New System.Drawing.Point(0, 0)
        Me.tbpnMenuSeperator.Name = "tbpnMenuSeperator"
        Me.tbpnMenuSeperator.RowCount = 1
        Me.tbpnMenuSeperator.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tbpnMenuSeperator.Size = New System.Drawing.Size(350, 37)
        Me.tbpnMenuSeperator.TabIndex = 0
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 2000
        '
        'MenuBar
        '
        Me.Controls.Add(Me.tbpnMenuSeperator)
        Me.Name = "MenuBar"
        Me.Size = New System.Drawing.Size(350, 37)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tbpnMenuSeperator As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
