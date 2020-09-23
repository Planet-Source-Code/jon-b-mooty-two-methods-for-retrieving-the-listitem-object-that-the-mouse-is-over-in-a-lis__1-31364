<div align="center">

## Two methods for retrieving the ListItem object that the mouse is over in a ListView control\.


</div>

### Description

This code gives you two methods (one using built-in functions of the ListView control, and one using SendMessage and the LVM_GETITEMRECT constant) to retrieve the ListItem object that the mouse pointer is currently over with. This is demonstrated by highlighting the item.
 
### More Info
 
To test the SendMessage method set the constant USE_SENDMESSAGE to True.

The current ListItem object the mouse pointer is over.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon B\. Mooty](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-b-mooty.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-b-mooty-two-methods-for-retrieving-the-listitem-object-that-the-mouse-is-over-in-a-lis__1-31364/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_GETITEMRECT = &H100E
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
```


### Source Code

```
' change this to true in order
' to test the SendMessage method
#Const USE_SENDMESSAGE = False
' you will need a ListView control
' named lvw in order to test the
' following code
Private Sub Form_Load()
  Dim iCol As Integer, iRow As Integer
  Dim sPos As String
  ' setup listview
  ' use report view with 3 columns and 500
  '   rows
  lvw.View = lvwReport
  lvw.MultiSelect = False
  lvw.ColumnHeaders.Add , , "Column1"
  lvw.ColumnHeaders.Add , , "Column2"
  lvw.ColumnHeaders.Add , , "Column3"
  For iRow = 1 To 500
    For iCol = 1 To 3
      sPos = "Row = " & iRow & ", " & "Col = " & iCol
      If iCol = 1 Then
        lvw.ListItems.Add , , sPos
      Else
        lvw.ListItems(iRow).SubItems(iCol - 1) = sPos
      End If
    Next iCol
  Next iRow
End Sub
Public Function GetListViewItemFromPt(plvw As ListView, x As Single, y As Single) As ListItem
  Dim bFound As Boolean
  Dim rCur As RECT
  Dim iIndex As Integer
  Dim iPixX As Integer, iPixY As Integer
  ' convert the coordinates to
  ' pixels (remove the next two
  ' lines if you will be passing
  ' the coordinates in pixels)
  iPixX = Me.ScaleX(x, vbTwips, vbPixels)
  iPixY = Me.ScaleY(y, vbTwips, vbPixels)
  For iIndex = 1 To plvw.ListItems.Count
    ' get the coordinates for each
    ' item in the listbox in its current
    ' state, if Top is less than 0 or Bottom
    '
    ' is greater than the height of the
    ' listbox then the item is currently out
    '   of the
    ' viewable area
    rCur.Left = 0
    SendMessage plvw.hwnd, LVM_GETITEMRECT, iIndex, rCur
    ' if passed corrdinates are within
    ' the bounds of the current item than
    ' exit the loop and return the ListItem
    '   object
    If iPixY >= rCur.Top And iPixY <= rCur.Bottom Then
      bFound = True
      Exit For
    End If
  Next iIndex
  If bFound Then
    Set GetListViewItemFromPt = plvw.ListItems(iIndex + 1)
  Else
    Set GetListViewItemFromPt = Nothing
  End If
End Function
Private Sub lvw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' set the selected ListItem object equal
  '   to the Item the mouse
  ' is currently over
  #If USE_SENDMESSAGE Then
    Set lvw.SelectedItem = GetListViewItemFromPt(lvw, x, y)
  #Else
    Set lvw.SelectedItem = lvw.HitTest(x, y)
  #End If
End Sub
```

