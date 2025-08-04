# Custom Enumerator/Collection Class
## Calling module
```vb
Public Sub Test()
    Dim Cats As Cats
    Set Cats = New Cats
    Cats.AddCat "foo"
    Cats.AddCat "bar"
    Cats.AddCat "baz"
    
    ' Works ok, no stack frame bug
    Dim ThisCat As Cat
    For Each ThisCat In Cats
        Debug.Print ThisCat.Name
    Next ThisCat
    
    ' Works OK. Note: Cats.Item(i) is of type Cat
    Dim i As Long
    For i = 1 To Cats.Count
        Debug.Print Cats.Item(i).Name
    Next i
End Sub
```
## Singular Class
```vb
Public Name As String
```

## Collection Class
```vb
Private Items As Collection

'@Enumerator
Public Property Get NewEnum() As IUnknown
    Set NewEnum = Items.[_NewEnum]
End Property

Public Property Get Count() As Long
    Count = Items.Count
End Property

' _Could_
Public Property Get Item(ByVal Index As Long) As Cat
    Set Item = Items.Item(Index)
End Property

Private Sub Class_Initialize()
    Set Items = New Collection
End Sub

Public Sub AddCat(ByVal Name As String)
    Dim NewCat As Cat
    Set NewCat = New Cat
    NewCat.Name = Name
    Items.Add NewCat
End Sub
```