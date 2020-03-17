Public Class frmWorkersBuffer
    Public WriteOnly Property HeaderMessage As String
        Set(value As String)
            lbHeader.Text = value
        End Set
    End Property

    Public WriteOnly Property RemainingProcesses As Integer
        Set(value As Integer)
            lbProcess.Text = String.Format("Remaining Process: {0} ...", value)
        End Set
    End Property

End Class