Attribute VB_Name = "mdldatabase"
Public rs As Recordset
Public db As Database
Public wks As Workspace
Public accounttype As String
Public username As String
Public password As String
Public tries As Integer


Public Function OpenDbase()
Dim dbname As String
    On Error GoTo NotOPen
    Set wks = DBEngine.Workspaces(0)
    dbname = App.Path & "\francissalvadorbcopassword.mdb"
    Set db = wks.OpenDatabase(dbname, False, False, ";PWD=francissalvadorbco")
    Set rs = db.OpenRecordset("password", dbOpenDynaset)
    
    Exit Function
NotOPen:
    MsgBox "Error" & Str$(Err.Number) & _
        " opening database " & dbname & "." & _
        vbCrLf & Err.Description
End Function

