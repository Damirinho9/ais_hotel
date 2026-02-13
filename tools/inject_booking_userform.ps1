[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$MacroModulePath = ".\vba\HotelMacros.bas",
    [string]$UserFormPath = ".\vba\frmBookingAdd.frm"
)

$ErrorActionPreference = 'Stop'

function Resolve-RelativePath {
    param([string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) { return (Resolve-Path $Path).Path }
    return (Resolve-Path (Join-Path $PSScriptRoot "..\$Path")).Path
}

function Get-FormCodeFromFrm {
    param([string]$FrmContent)
    $start = $FrmContent.IndexOf("Private Sub UserForm_Initialize()")
    if ($start -lt 0) {
        throw "Не найден код формы (Private Sub UserForm_Initialize) в файле $UserFormPath"
    }
    return $FrmContent.Substring($start)
}

function Ensure-VBComponent {
    param(
        $VBProject,
        [string]$Name,
        [int]$Type
    )

    foreach ($comp in $VBProject.VBComponents) {
        if ($comp.Name -eq $Name) { return $comp }
    }

    $newComp = $VBProject.VBComponents.Add($Type)
    $newComp.Name = $Name
    return $newComp
}

function Replace-Code {
    param(
        $CodeModule,
        [string]$CodeText
    )
    $lineCount = $CodeModule.CountOfLines
    if ($lineCount -gt 0) {
        $CodeModule.DeleteLines(1, $lineCount)
    }
    $CodeModule.AddFromString($CodeText)
}

function Add-Control {
    param(
        $Designer,
        [string]$ProgId,
        [string]$Name,
        [int]$Left,
        [int]$Top,
        [int]$Width,
        [int]$Height,
        [string]$Caption = $null
    )

    try {
        $existing = $Designer.Controls.Item($Name)
        if ($null -ne $existing) {
            $existing.Left = $Left
            $existing.Top = $Top
            $existing.Width = $Width
            $existing.Height = $Height
            if ($null -ne $Caption) { $existing.Caption = $Caption }
            return $existing
        }
    }
    catch {}

    $ctrl = $Designer.Controls.Add($ProgId, $Name, $true)
    $ctrl.Left = $Left
    $ctrl.Top = $Top
    $ctrl.Width = $Width
    $ctrl.Height = $Height
    if ($null -ne $Caption) { $ctrl.Caption = $Caption }
    return $ctrl
}

$workbookFullPath = if ([System.IO.Path]::IsPathRooted($WorkbookPath)) { $WorkbookPath } else { Join-Path (Get-Location) $WorkbookPath }
if (-not (Test-Path $workbookFullPath)) {
    throw "Файл книги не найден: $workbookFullPath"
}

$macroModuleFullPath = Resolve-RelativePath $MacroModulePath
$userFormFullPath = Resolve-RelativePath $UserFormPath

$macroCode = Get-Content -LiteralPath $macroModuleFullPath -Raw -Encoding UTF8
$frmRaw = Get-Content -LiteralPath $userFormFullPath -Raw -Encoding UTF8
$formCode = Get-FormCodeFromFrm -FrmContent $frmRaw

$xl = $null
$wb = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.DisplayAlerts = $false
    $xl.Visible = $false

    $wb = $xl.Workbooks.Open($workbookFullPath)
    $vbProj = $wb.VBProject

    # 1 = vbext_ct_StdModule
    $moduleComp = Ensure-VBComponent -VBProject $vbProj -Name "HotelMacros" -Type 1
    Replace-Code -CodeModule $moduleComp.CodeModule -CodeText $macroCode

    # 3 = vbext_ct_MSForm
    $formComp = Ensure-VBComponent -VBProject $vbProj -Name "frmBookingAdd" -Type 3
    $designer = $formComp.Designer

    $designer.Caption = "Добавить бронирование"
    $designer.Width = 330
    $designer.Height = 305

    # Labels
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblRoom" -Left 12 -Top 18 -Width 85 -Height 18 -Caption "№ Комнаты:" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblGuest" -Left 12 -Top 52 -Width 85 -Height 18 -Caption "Гость:" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblCheckIn" -Left 12 -Top 86 -Width 85 -Height 18 -Caption "Дата заезда:" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblCheckOut" -Left 12 -Top 120 -Width 85 -Height 18 -Caption "Дата выезда:" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblGuestsCount" -Left 12 -Top 154 -Width 85 -Height 18 -Caption "Кол-во гостей:" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblPrice" -Left 12 -Top 188 -Width 85 -Height 18 -Caption "Цена/сутки:" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.Label.1" -Name "lblStatus" -Left 12 -Top 222 -Width 85 -Height 18 -Caption "Статус:" | Out-Null

    # Inputs
    Add-Control -Designer $designer -ProgId "Forms.ComboBox.1" -Name "cboRoom" -Left 104 -Top 14 -Width 200 -Height 20 | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.ComboBox.1" -Name "cboGuest" -Left 104 -Top 48 -Width 200 -Height 20 | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.TextBox.1" -Name "txtCheckIn" -Left 104 -Top 82 -Width 200 -Height 20 | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.TextBox.1" -Name "txtCheckOut" -Left 104 -Top 116 -Width 200 -Height 20 | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.TextBox.1" -Name "txtGuestsCount" -Left 104 -Top 150 -Width 200 -Height 20 | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.TextBox.1" -Name "txtPrice" -Left 104 -Top 184 -Width 200 -Height 20 | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.ComboBox.1" -Name "cboStatus" -Left 104 -Top 218 -Width 200 -Height 20 | Out-Null

    # Buttons
    Add-Control -Designer $designer -ProgId "Forms.CommandButton.1" -Name "btnSave" -Left 104 -Top 252 -Width 96 -Height 26 -Caption "Сохранить" | Out-Null
    Add-Control -Designer $designer -ProgId "Forms.CommandButton.1" -Name "btnCancel" -Left 208 -Top 252 -Width 96 -Height 26 -Caption "Отмена" | Out-Null

    Replace-Code -CodeModule $formComp.CodeModule -CodeText $formCode

    # Привязка кнопки на листе Бронирование
    try {
        $ws = $wb.Worksheets("Бронирование")
        try {
            $shape = $ws.Shapes.Item("btn_add_booking")
            if ($null -ne $shape) {
                $shape.OnAction = "ДобавитьБронь"
            }
        }
        catch {}

        try {
            $ole = $ws.OLEObjects("btn_add_booking")
            if ($null -ne $ole) {
                $ole.Object.Caption = "Добавить бронирование"
            }
        }
        catch {}
    }
    catch {
        Write-Warning "Лист Бронирование или кнопка btn_add_booking не найдены для автопривязки."
    }

    $wb.Save()
    Write-Host "Готово: форма и макрос внедрены в $workbookFullPath"
}
finally {
    if ($null -ne $wb) { $wb.Close($true) | Out-Null }
    if ($null -ne $xl) {
        $xl.Quit()
        if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
