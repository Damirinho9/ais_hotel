# Инструкция по внедрению формы бронирования в `АИС_Гостиница.xlsm`

Добавлены исходники VBA:
- `vba/frmBookingAdd.frm` — код UserForm `frmBookingAdd`.
- `vba/HotelMacros.bas` — модуль с процедурой `ДобавитьБронь`, открывающей форму.
- `tools/inject_booking_userform.ps1` — автоматическая инъекция формы/макросов в `.xlsm` через Excel COM.

## Быстрый способ (рекомендуется)
> Требуется **Windows + установленный Excel** и включённый доступ к VBA Project:
> Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → **Trust access to the VBA project object model**.

### Запуск из PowerShell
Если вы **уже внутри PowerShell**, запускайте так (без `-ExecutionPolicy` в начале строки):

```powershell
.\tools\inject_booking_userform.ps1 -WorkbookPath ".\АИС_Гостиница.xlsm"
```

Если политика выполнения блокирует скрипт, выполните:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\tools\inject_booking_userform.ps1 -WorkbookPath ".\АИС_Гостиница.xlsm"
```

Либо запустите команду **из cmd/Win+R**, где допустим префикс `powershell`:

```cmd
powershell -ExecutionPolicy Bypass -File .\tools\inject_booking_userform.ps1 -WorkbookPath ".\АИС_Гостиница.xlsm"
```

## Ручной способ (если автозапуск недоступен)
1. Откройте `АИС_Гостиница.xlsm` в Excel.
2. Нажмите `Alt+F11`.
3. `File -> Import File...` и импортируйте `vba/HotelMacros.bas`.
4. Создайте `Insert -> UserForm` и в окне Properties задайте имя `frmBookingAdd`.
5. Добавьте на форму контролы с именами:
   - `cboRoom` (ComboBox)
   - `cboGuest` (ComboBox)
   - `txtCheckIn` (TextBox)
   - `txtCheckOut` (TextBox)
   - `txtGuestsCount` (TextBox)
   - `txtPrice` (TextBox)
   - `cboStatus` (ComboBox)
   - `btnSave` (CommandButton, Caption: `Сохранить`)
   - `btnCancel` (CommandButton, Caption: `Отмена`)
6. Откройте Code у `frmBookingAdd` и вставьте код из `vba/frmBookingAdd.frm` начиная с `Private Sub UserForm_Initialize()`.
7. Проверьте привязку кнопки `btn_add_booking` на листе `Бронирование` к макросу `ДобавитьБронь` (`Assign Macro...`).
8. Сохраните файл в формате `.xlsm`.

## Ручная проверка
1. Нажмите «Добавить бронирование».
2. Заполните форму.
3. Нажмите «Сохранить».
4. Убедитесь, что строка появилась в таблице листа `Бронирование`.
