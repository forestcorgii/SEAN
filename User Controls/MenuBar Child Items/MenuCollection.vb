Imports System.ComponentModel

Public Class MenuCollection
    Inherits System.Windows.Forms.BaseCollection
    Implements ICollection, IList, IEnumerable

    'Public Shadows Sub Add(_menu As MenuButton)
    '    _menu.Name = "Menu" & Count + 1
    '    List.Add(_menu)
    'End Sub

    'Public Shadows Sub Add(_menu As String)
    '    Dim m As New MenuButton(_menu)
    '    List.Add(m)
    'End Sub

    Public Function Length() As Integer
        Return Count
    End Function


    Public Function Add(value As Object) As Integer Implements IList.Add
        Return List.Add(value)
    End Function

    Public Sub Clear() Implements IList.Clear
        List.Clear()
    End Sub

    Public Function Contains(value As Object) As Boolean Implements IList.Contains
        Return List.Contains(value)
    End Function

    Public Function IndexOf(value As Object) As Integer Implements IList.IndexOf
        Return List.IndexOf(value)
    End Function

    Public Sub Insert(index As Integer, value As Object) Implements IList.Insert
        List.Insert(index, value)
    End Sub

    Public ReadOnly Property IsFixedSize As Boolean Implements IList.IsFixedSize
        Get
            Return List.IsFixedSize
        End Get
    End Property

    Public ReadOnly Property IsReadOnly1 As Boolean Implements IList.IsReadOnly
        Get
            Return List.IsReadOnly
        End Get
    End Property

    Default Public Overloads Property Item(index As Integer) As Object Implements IList.Item
        Get
            Return List.Item(index)
        End Get
        Set(value As Object)
            List.Item(index) = value
        End Set
    End Property

    Public Sub Remove(value As Object) Implements IList.Remove
        List.Remove(value)
    End Sub

    Public Sub RemoveAt(index As Integer) Implements IList.RemoveAt
        List.RemoveAt(index)
    End Sub
End Class
