# Инструкция по внедрению формы бронирования в `АИС_Гостиница.xlsm`

Добавлены исходники VBA:
- `vba/frmBookingAdd.frm` — код UserForm `frmBookingAdd`.
- `vba/HotelMacros.bas` — модуль с процедурой `ДобавитьБронь`, открывающей форму.

## Что нужно сделать в Excel
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
6. Откройте Code у `frmBookingAdd` и вставьте код из `vba/frmBookingAdd.frm` начиная с `Private Sub UserForm_Initialize()` (блок `Begin VB.UserForm ... End` относится к экспортированному формату и может не вставляться вручную).
7. Проверьте привязку кнопки `btn_add_booking` на листе `Бронирование` к макросу `ДобавитьБронь` (`Assign Macro...`).
8. Сохраните файл в формате `.xlsm`.

## Ручная проверка
1. Нажмите «Добавить бронирование».
2. Заполните форму.
3. Нажмите «Сохранить».
4. Убедитесь, что строка появилась в таблице листа `Бронирование`.
