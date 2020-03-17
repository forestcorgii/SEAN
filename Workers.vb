Imports System.ComponentModel
Imports System.Threading
Imports System.Windows.Forms
'Public Class SeanWorker
'    Inherits BackgroundWorker
'    Public Shadows Event RunWorkerCompleted(sender As Object, e As SeanWorkerCompletedEventArgs)
'    Public Shadows Event DoWork(sender As Object, e As DoWorkNowEventArgs)
'    Public _IsBusy As Boolean
'    ' Public 

'    Public Class DoWorkNowEventArgs
'        Inherits CancelEventArgs
'        Public argument As Object
'        Public cancel As Boolean
'        Public result As Object
'    End Class
'    Public Class SeanWorkerCompletedEventArgs
'        Inherits CancelEventArgs
'        Public argument As Object
'        Public cancel As Boolean
'        Public result As Object
'    End Class
'    Public Shadows Property IsBusy As Boolean
'        Get
'            Return _IsBusy
'        End Get
'        Set(value As Boolean)
'            _IsBusy = value
'        End Set
'    End Property


'    Private proc As Integer = 0
'    Public Shadows Sub RunWorkerAsync(args As Object)
'        proc = 0
'        IsBusy = True
'        Dim t As New Thread(New ThreadStart(AddressOf DoWorksub))
'        t.Start()

'        IsBusy = False
'    End Sub
'    Private Function DoWorksub()
'        RaiseEvent DoWork(Me, Nothing)
'    End Function
'End Class

Public Class Workers
    Private workers As List(Of BackgroundWorker)
    Public Event Worker_DoWork(sender As Object, e As DoWorkEventArgs)
    Public Event Worker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
    Public Event RunWorkersCompleted(sender As Object)
    Public Event WorkersWorkBegin(sender As Object)

    Public QueueArgs As List(Of Object)
    Public UseBuffer As Boolean
    Private Buffer As frmWorkersBuffer
    Private mainForm As Form
#Region "Properties"
    Public ReadOnly Property IsBusy As Boolean
        Get
            Return busyWorkersCount > 0
        End Get
    End Property

    Private ReadOnly Property busyWorkersCount As Integer
        Get
            Return (From res As BackgroundWorker In workers Where res.IsBusy Select res).ToList.Count
        End Get
    End Property
#End Region

    Sub New(Optional _mainForm As Form = Nothing, Optional queueInterval As Integer = 100, Optional _useBuffer As Boolean = False)
        mainForm = _mainForm
        Buffer = New frmWorkersBuffer
        QueueArgs = New List(Of Object)
        init_Timer()
        SetTimer(queueInterval)
        SetWorker(3)
        UseBuffer = _useBuffer
    End Sub

    Public Sub SetWorker(workerCount As Integer)
        workers = New List(Of BackgroundWorker)
        For i As Integer = 0 To workerCount - 1
            Dim newWorker As New BackgroundWorker
            With newWorker
                AddHandler .DoWork, AddressOf Item_DoWork
                AddHandler .RunWorkerCompleted, AddressOf Item_RunWorkerCompleted
            End With
            workers.Add(newWorker)
        Next
    End Sub

#Region "Event Handlers"
    Private Sub Item_DoWork(sender As Object, e As DoWorkEventArgs)
        RaiseEvent Worker_DoWork(sender, e)
    End Sub

    Private Sub Item_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        RaiseEvent Worker_RunWorkerCompleted(sender, e)
    End Sub
#End Region

    Public Sub AddtoQueue(arg As Object)
        QueueArgs.Add(arg)
        If Not ticking Then
            StartWorking()
            showBuffer()
        End If
    End Sub

    Public Sub AddRangetoQueue(args As List(Of Object))
        QueueArgs.AddRange(args)
        If Not ticking Then
            StartWorking()
            showBuffer()
        End If
    End Sub

    Public Function AttemptExecute(Optional args As Object = Nothing) As Boolean
        For i As Integer = 0 To workers.Count - 1
            If workers(i).IsBusy = False Then
                workers(i).RunWorkerAsync(args)
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub UpdateQueue()
        If QueueArgs.Count > 0 Then
            If AttemptExecute(QueueArgs(0)) Then
                QueueArgs.RemoveAt(0)
            End If
        End If
    End Sub

#Region "Timer"
    Private queueTimer As Threading.Timer
    Private interval As Integer = -1
    Private ticking As Boolean

    Private Sub init_Timer()
        queueTimer = New Threading.Timer(AddressOf queueTimer_Tick, Nothing, Timeout.Infinite, Timeout.Infinite)
    End Sub

    Private Sub queueTimer_Tick(sender As Object)
        UpdateQueue()
        updateBuffer()
        If Not IsBusy Then
            StartWorking()
            hideBuffer()
            RaiseEvent RunWorkersCompleted(Me)
        End If
    End Sub

    Public Sub StartWorking()
        queueTimer.Change(0, interval)
        ticking = True
    End Sub

    Public Sub StopWorking()
        queueTimer.Change(-1, -1)
        ticking = False
    End Sub

    Private Sub SetTimer(_interval As Integer)
        interval = _interval
    End Sub
#End Region

#Region "Buffer"
    Private Sub showBuffer()
        If UseBuffer And mainForm IsNot Nothing Then
            mainForm.Invoke(Sub() Buffer.Show())
        End If
    End Sub
    Private Sub hideBuffer()
        If UseBuffer And mainForm IsNot Nothing Then
            mainForm.Invoke(Sub() Buffer.Hide())
        End If
    End Sub
    Private Sub updateBuffer()
        If UseBuffer And mainForm IsNot Nothing Then
            mainForm.Invoke(Sub() Buffer.RemainingProcesses = busyWorkersCount + QueueArgs.Count)
        End If
    End Sub
#End Region

End Class
