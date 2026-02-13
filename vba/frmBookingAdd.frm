VERSION 5.00
Begin VB.UserForm frmBookingAdd
   Caption         =   "Добавить бронирование"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBookingAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    cboRoom.Clear
    cboGuest.Clear
    cboStatus.Clear

    FillComboFromColumn cboRoom, "НомернойФонд", "№ Комнаты"
    FillComboFromColumn cboGuest, "Гости", "ФИО"

    cboStatus.AddItem "Бронь"
    cboStatus.AddItem "Активна"
    cboStatus.AddItem "Завершена"

    txtCheckIn.Value = Format(Date, "dd.mm.yyyy")
    txtCheckOut.Value = Format(Date + 1, "dd.mm.yyyy")
    txtGuestsCount.Value = "1"
End Sub

Private Sub btnCancel_Click()
    ClearForm
    Unload Me
End Sub

Private Sub btnSave_Click()
    If Not ValidateForm Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Бронирование")

    Dim newRow As Long
    Dim lo As ListObject

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        lo.ListRows.Add
        newRow = lo.ListRows(lo.ListRows.Count).Range.Row
    Else
        newRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    End If

    Dim checkInDate As Date
    Dim checkOutDate As Date
    checkInDate = CDate(txtCheckIn.Value)
    checkOutDate = CDate(txtCheckOut.Value)

    Dim nights As Long
    nights = DateDiff("d", checkInDate, checkOutDate)

    ws.Cells(newRow, "C").Value = GetNextBookingId(ws)
    ws.Cells(newRow, "D").Value = cboRoom.Value
    ws.Cells(newRow, "F").Value = cboGuest.Value
    ws.Cells(newRow, "G").Value = checkInDate
    ws.Cells(newRow, "H").Value = checkOutDate
    ws.Cells(newRow, "I").Value = nights
    ws.Cells(newRow, "L").Value = cboStatus.Value
    ws.Cells(newRow, "M").Value = CLng(txtGuestsCount.Value)

    If IsNumeric(txtPrice.Value) Then
        ws.Cells(newRow, "J").Value = CDbl(txtPrice.Value)
        ws.Cells(newRow, "K").Value = nights * CDbl(txtPrice.Value)
    End If

    MsgBox "Бронирование успешно добавлено.", vbInformation
    ClearForm
    Unload Me
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(cboRoom.Value) = "" Then
        MsgBox "Укажите номер комнаты.", vbExclamation
        cboRoom.SetFocus
        Exit Function
    End If

    If Trim(cboGuest.Value) = "" Then
        MsgBox "Укажите гостя.", vbExclamation
        cboGuest.SetFocus
        Exit Function
    End If

    If Not IsDate(txtCheckIn.Value) Then
        MsgBox "Введите корректную дату заезда.", vbExclamation
        txtCheckIn.SetFocus
        Exit Function
    End If

    If Not IsDate(txtCheckOut.Value) Then
        MsgBox "Введите корректную дату выезда.", vbExclamation
        txtCheckOut.SetFocus
        Exit Function
    End If

    If CDate(txtCheckOut.Value) < CDate(txtCheckIn.Value) Then
        MsgBox "Дата выезда не может быть раньше даты заезда.", vbExclamation
        txtCheckOut.SetFocus
        Exit Function
    End If

    If Trim(cboStatus.Value) = "" Then
        MsgBox "Укажите статус брони.", vbExclamation
        cboStatus.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtGuestsCount.Value) Or CLng(txtGuestsCount.Value) <= 0 Then
        MsgBox "Количество гостей должно быть положительным числом.", vbExclamation
        txtGuestsCount.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

Private Sub ClearForm()
    cboRoom.Value = ""
    cboGuest.Value = ""
    txtCheckIn.Value = ""
    txtCheckOut.Value = ""
    txtGuestsCount.Value = ""
    txtPrice.Value = ""
    cboStatus.Value = ""
End Sub

Private Function GetNextBookingId(ByVal ws As Worksheet) As String
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    Dim lastId As String
    Dim nextNum As Long

    If lastRow < 11 Then
        nextNum = 1
    Else
        lastId = CStr(ws.Cells(lastRow, "C").Value)
        nextNum = Val(Replace(lastId, "Б", "")) + 1
        If nextNum <= 0 Then nextNum = lastRow - 9
    End If

    GetNextBookingId = "Б" & Format(nextNum, "000")
End Function

Private Sub FillComboFromColumn(ByVal cbo As MSForms.ComboBox, ByVal sheetName As String, ByVal headerName As String)
    On Error GoTo SafeExit

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim headerCell As Range
    Set headerCell = ws.Rows(10).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If headerCell Is Nothing Then Set headerCell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If headerCell Is Nothing Then GoTo SafeExit

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, headerCell.Column).End(xlUp).Row

    Dim i As Long
    For i = headerCell.Row + 1 To lastRow
        If Trim(CStr(ws.Cells(i, headerCell.Column).Value)) <> "" Then
            cbo.AddItem CStr(ws.Cells(i, headerCell.Column).Value)
        End If
    Next i

SafeExit:
End Sub
