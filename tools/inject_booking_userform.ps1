param(
  [Parameter(Mandatory=$true)] [string]$WorkbookPath
)

$ErrorActionPreference = 'Stop'

if (!(Test-Path $WorkbookPath)) {
  throw "Workbook not found: $WorkbookPath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path)
  $vbProj = $wb.VBProject

  # Ensure module HotelMacros exists
  $stdModule = $null
  foreach ($c in $vbProj.VBComponents) {
    if ($c.Type -eq 1 -and $c.Name -eq 'HotelMacros') { $stdModule = $c; break }
  }
  if ($null -eq $stdModule) {
    $stdModule = $vbProj.VBComponents.Add(1)
    $stdModule.Name = 'HotelMacros'
  }

  $macroCode = @'
Attribute VB_Name = "HotelMacros"
Option Explicit

Public Sub ДобавитьБронь()
    frmBookingAdd.Show
End Sub
'@

  $stdModule.CodeModule.DeleteLines(1, $stdModule.CodeModule.CountOfLines)
  $stdModule.CodeModule.AddFromString($macroCode)

  # Create or reuse UserForm
  $frm = $null
  foreach ($c in $vbProj.VBComponents) {
    if ($c.Type -eq 3 -and $c.Name -eq 'frmBookingAdd') { $frm = $c; break }
  }
  if ($null -eq $frm) {
    $frm = $vbProj.VBComponents.Add(3)
    $frm.Name = 'frmBookingAdd'
  }

  $designer = $frm.Designer
  $designer.Caption = 'Добавить бронирование'
  $designer.Width = 420
  $designer.Height = 330

  # Remove old controls
  for ($i = $designer.Controls.Count; $i -ge 1; $i--) {
    $designer.Controls.Remove($designer.Controls.Item($i-1).Name)
  }

  function Add-Label($name, $caption, $left, $top) {
    $c = $designer.Controls.Add('Forms.Label.1', $name, $true)
    $c.Caption = $caption
    $c.Left = $left
    $c.Top = $top
    $c.Width = 140
    return $c
  }

  function Add-Text($name, $left, $top) {
    $c = $designer.Controls.Add('Forms.TextBox.1', $name, $true)
    $c.Left = $left
    $c.Top = $top
    $c.Width = 220
    return $c
  }

  Add-Label 'lblGuest' 'Гость' 16 18 | Out-Null
  Add-Text 'txtGuest' 160 16 | Out-Null

  Add-Label 'lblRoom' 'Номер' 16 50 | Out-Null
  Add-Text 'txtRoom' 160 48 | Out-Null

  Add-Label 'lblCheckIn' 'Дата заезда (дд.мм.гггг)' 16 82 | Out-Null
  Add-Text 'txtCheckIn' 160 80 | Out-Null

  Add-Label 'lblCheckOut' 'Дата выезда (дд.мм.гггг)' 16 114 | Out-Null
  Add-Text 'txtCheckOut' 160 112 | Out-Null

  Add-Label 'lblStatus' 'Статус' 16 146 | Out-Null
  Add-Text 'txtStatus' 160 144 | Out-Null

  Add-Label 'lblPayment' 'Оплата' 16 178 | Out-Null
  Add-Text 'txtPayment' 160 176 | Out-Null

  $btnSave = $designer.Controls.Add('Forms.CommandButton.1', 'btnSave', $true)
  $btnSave.Caption = 'Сохранить'
  $btnSave.Left = 160
  $btnSave.Top = 230
  $btnSave.Width = 100

  $btnCancel = $designer.Controls.Add('Forms.CommandButton.1', 'btnCancel', $true)
  $btnCancel.Caption = 'Отмена'
  $btnCancel.Left = 280
  $btnCancel.Top = 230
  $btnCancel.Width = 100

  $formCode = @'
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim ws As Worksheet
    Dim nextRow As Long

    If Trim$(txtGuest.Value) = "" Or Trim$(txtRoom.Value) = "" Then
        MsgBox "Заполните обязательные поля: Гость и Номер.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(txtCheckIn.Value) Or Not IsDate(txtCheckOut.Value) Then
        MsgBox "Введите корректные даты заезда и выезда.", vbExclamation
        Exit Sub
    End If

    If CDate(txtCheckOut.Value) < CDate(txtCheckIn.Value) Then
        MsgBox "Дата выезда не может быть раньше даты заезда.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("Бронирование")
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = nextRow - 1
    ws.Cells(nextRow, 2).Value = Trim$(txtGuest.Value)
    ws.Cells(nextRow, 3).Value = Trim$(txtRoom.Value)
    ws.Cells(nextRow, 4).Value = CDate(txtCheckIn.Value)
    ws.Cells(nextRow, 5).Value = CDate(txtCheckOut.Value)
    ws.Cells(nextRow, 6).Value = Trim$(txtStatus.Value)
    ws.Cells(nextRow, 7).Value = Trim$(txtPayment.Value)

    MsgBox "Бронирование добавлено.", vbInformation
    Unload Me
End Sub
'@

  $frm.CodeModule.DeleteLines(1, $frm.CodeModule.CountOfLines)
  $frm.CodeModule.AddFromString($formCode)

  # Re-assign shape macro if shape exists
  $ws = $wb.Worksheets.Item('Бронирование')
  foreach ($shape in $ws.Shapes) {
    if ($shape.Name -eq 'btn_add_booking') {
      $shape.OnAction = 'ДобавитьБронь'
    }
  }

  $wb.Save()
  $wb.Close($true)
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Done: UserForm embedded into $WorkbookPath"
