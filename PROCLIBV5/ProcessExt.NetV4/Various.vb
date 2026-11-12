Public Class Collection_Class
    Dim C As Collection
    Public Function Add(ByVal Item As Object, ByVal Key As Object) As Boolean
        C.Add(Item, Key)
    End Function

    Default Public ReadOnly Property Item(ByVal Key As Object) As Object
        Get
            On Error GoTo errmsg
            Item = C(Key)
            Exit Property
errmsg:
            Item = ""
        End Get
    End Property
    Public ReadOnly Property Count() As Integer
        Get
            On Error GoTo errmsg
            Count = C.Count
            Exit Property
errmsg:
            Count = 0
        End Get
    End Property
    Public Sub Clear()
        Dim i As Integer
        Dim j As Integer
        j = C.Count
        For i = 1 To j
            C.Remove(1)
        Next
    End Sub

    Public Sub New()
        C = New Collection
    End Sub
End Class
