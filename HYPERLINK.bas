Attribute VB_Name = "HYPERLINK"
Option Explicit

Sub HYPERLINK()
Attribute HYPERLINK.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim adopt As Worksheet
    Dim adptlist As ListObject
    Dim i As Integer
    Set adopt = Workbooks("FA23 BUYING - ADOPTION LIST.xlsm").Worksheets("ADOPTION LIST")
    Set adptlist = adopt.ListObjects("ADOPTION_LIST")
    
    For i = 1 To adptlist.ListRows.Count
        If adptlist.DataBodyRange(i, 4).Value <> "" Then
            adptlist.DataBodyRange(i, 3).Hyperlinks.add _
                Anchor:=adptlist.DataBodyRange(i, 3), _
                Address:=adptlist.DataBodyRange(i, 4).Value, _
                TextToDisplay:=adptlist.DataBodyRange(i, 3).Value
        End If
    Next i
End Sub
