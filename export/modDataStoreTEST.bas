Attribute VB_Name = "modDataStoreTEST"
'@Folder("RemoteDataStore")
Option Explicit

'@Ignore ProcedureNotUsed
Private Sub ZZZTESTRemoteTables()
    Dim Model As DataStore
    Set Model = New DataStore
    Model.Load
    Debug.Print Model.Remote.Tables.Count
        
    Dim SerStr As String
    SerStr = "MyKey¤Test\NATO§DeletionTime¤575E¥CreationTime¤324F¥ModifiedTime¤A2AC"
    Model.Remote.Tables.Add SerStr, "Catland", "1234"
    Model.Remote.Tables.Save
End Sub

'@Ignore ProcedureNotUsed
Private Sub ZZZTESTRemoteMaps()
    Dim Model As DataStore
    Set Model = New DataStore
    Model.Load
    Debug.Print Model.Remote.Maps.Count
        
    Dim SerStr As String
    SerStr = "MyKey¤Test\NATO§DeletionTime¤575E¥CreationTime¤324F¥ModifiedTime¤A2AC"
    Model.Remote.Maps.Add SerStr, "Catland"
    Model.Remote.Maps.Save
End Sub

'@Ignore ProcedureNotUsed
Private Sub ZZZTESTRemoteGetKeyIDs()
    Dim Model As DataStore
    Set Model = New DataStore
    Model.Load
    
    Dim foo As Variant
    ReDim foo(1 To 3)
    foo(1) = "foo2"
    foo(2) = "bar3"
    foo(3) = "baz"
    
    Dim Result As Variant
    Result = Model.Remote.Keys.GetIDsFromQuery("Test\PATO", foo)
End Sub

'@Ignore ProcedureNotUsed
Private Sub ZZZTESTRemoteGetValues()
    Dim Model As DataStore
    Set Model = New DataStore
    Model.Load
    
    Dim KeyIDs(1 To 3) As Variant
    KeyIDs(1) = "067C"
    KeyIDs(2) = "2B28"
    KeyIDs(3) = "44FD"
    
    Dim FieldIDs(1 To 3) As Variant
    FieldIDs(1) = "324F"
    FieldIDs(2) = "575E"
    FieldIDs(3) = "A2AC"
    
    Dim Existing As Variant
    ReDim Existing(1 To 3, 1 To 3)
    Existing(1, 3) = "bat!"
    
    Dim ChangeMask As Variant
    ChangeMask = Model.Remote.Values.GetChangeMask(KeyIDs, FieldIDs, Existing)
    
    Dim Result As Variant
    Result = Model.Remote.Values.GetValues(KeyIDs, FieldIDs)
    
    'Result(1, 3) = "bat!"
    
    'Model.Remote.Push KeyIDs, FieldIDs, Result
End Sub

