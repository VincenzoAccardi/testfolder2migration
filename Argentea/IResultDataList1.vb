Imports System.Collections.Generic

Protected Interface IResultDataList1(Of T)
    ReadOnly Property Count As Integer
    Default Property Item(index As Integer) As T
    Sub Add(item As T)
    Sub Clear()
    Sub CopyTo(array() As T, arrayIndex As Integer)
    Sub Insert(index As Integer, item As T)
    Sub RemoveAt(index As Integer)
    Function Contains(item As T) As Boolean
    Function GetEnumerator() As IEnumerator(Of T)
    Function IndexOf(item As T) As Integer
    Function NewPaid(BarCode As String, TerminalId As String) As PaidEntry
    Function Remove(item As T) As Boolean
End Interface
