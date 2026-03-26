# AutoCAD_Info_Panel

ObjectARX-плагин для AutoCAD 2007, который создаёт отдельную панель-окно и выводит в ней список параметров DWG.

## Что реализовано

- Команда `SHOWDWGPROPS` — открывает/показывает панель `DWG Properties`.
- Команда `HIDEDWGPROPS` — скрывает панель.
- На форме есть кнопка `Load params from XLSX`, которая запускает команду `XLSX2DWGPROP` (реализована на C++ ObjectARX).
- Кнопка `Insert selected custom FIELD` вставляет в TEXT/MTEXT объект-поле (field code) для выбранного CUSTOM-свойства: сначала пытается вставить напрямую в редактор (`EM_REPLACESEL`/`WM_PASTE`), а при неудаче оставляет поле в буфере обмена для вставки `Ctrl+V`. Кнопка активна и при редактировании существующего текста.
- При активном AutoCAD панель автоматически удерживается поверх других окон; при переключении на другое приложение режим topmost снимается.
- На панели отображаются:
  - Название параметра отображается увеличенным и жирным шрифтом, значение — обычным шрифтом.
  - системные переменные: `DWGNAME`, `DWGPREFIX`, `CTAB`, `CDATE`, `DWGTITLED`, `LOGINNAME`, `DBMOD`;
  - пользовательские свойства из окна `_DWGPROPS` (вкладка «Прочие») через `AcDbDatabaseSummaryInfo::getCustomSummaryInfo`.

## Файлы

- `src/DwgPropsPanelApp.cpp` — регистрация команд и интеграция с ARX lifecycle.
- `src/DwgPropsPanel.h`
- `src/DwgPropsPanel.cpp` — Win32-реализация окна панели и чтение свойств через `acedGetVar`.
- `src/Xlsx2DwgProp.cpp` — реализация команды `XLSX2DWGPROP` (Excel COM -> DWG properties).
- `src/ToolbarSetup.cpp` / `src/ToolbarSetup.h` — создание toolbar `MG-Panel` и кнопок `SHOW/HIDE`.
- `MG-Project.cui` — базовый файл MENUGROUP `MG-PROJECT` с преднастроенной toolbar `MG-Panel`.
- `MG-Project.mnu` — альтернативный текстовый menu-файл для MENUGROUP `MG-PROJECT` (legacy формат).
- `DwgPropsPanel.ini` — создается автоматически в пользовательском профиле (`%APPDATA%\\AutoCAD_Info_Panel\\DwgPropsPanel.ini`) для сохранения положения и размера панели.
- `Xlsx2DwgProp.ini` — настройки импорта (имя листа и номера колонок).

В `MG-Project.cui` и `MG-Project.mnu` предопределены кнопки панели `MG-Panel`:
- `SHOW` → `SHOWDWGPROPS`
- `HIDE` → `HIDEDWGPROPS`
Файл `MG-Project.mnu` оформлен в классическом формате toolbar (`_Toolbar` / `_Button`) с привязкой BMP-иконок `Loadpanel16/32.bmp` и `UnLoadpanel16/32.bmp`.


## Миграция на AutoCAD 2007 (выполнено в проекте)

- Проект `AutoCAD_Info_Pane.vcproj` переведён на линковку с библиотеками ObjectARX 2007 (`acdb17/acge17/...`).
- Жёстко заданные локальные пути к SDK 2006 заменены на переменные окружения (`ARX2007_SDK_INC`, `ARX2007_SDK_LIB`, `ARX2007_ACAD_DIR`), чтобы сборка была переносимой между машинами.
- Исходный код оставлен совместимым с текущей архитектурой плагина; дальнейшее развитие можно делать без повторной ручной миграции проекта.

## Важно для интеграции в существующий ARX-проект

Этот модуль **не содержит** собственного `acrxEntryPoint`.

В вашем `acrxEntryPoint.cpp` (класс `AcRxArxApp`) нужно вызвать:

- `DwgPropsPanel_Init(pkt)` внутри `On_kInitAppMsg`;
- `DwgPropsPanel_Unload()` внутри `On_kUnloadAppMsg`.

Пример:

```cpp
virtual AcRx::AppRetCode On_kInitAppMsg(void *pkt) {
    AcRx::AppRetCode retCode = AcRxArxApp::On_kInitAppMsg(pkt);
    DwgPropsPanel_Init(pkt);
    return retCode;
}

virtual AcRx::AppRetCode On_kUnloadAppMsg(void *pkt) {
    DwgPropsPanel_Unload();
    AcRx::AppRetCode retCode = AcRxArxApp::On_kUnloadAppMsg(pkt);
    return retCode;
}
```

## Подключение

1. Откройте `AutoCAD_Info_Pane.sln` в Visual Studio 2005.
2. В *Tools → Options → Projects and Solutions → VC++ Directories* (или через переменные окружения) задайте пути:
   - `ARX2007_SDK_INC` → `<ObjectARX 2007 SDK>\inc`;
   - `ARX2007_SDK_LIB` → `<ObjectARX 2007 SDK>\lib`;
   - `ARX2007_ACAD_DIR` → каталог установки AutoCAD 2007.
3. Проверьте, что в Linker используются библиотеки версии 2007: `acdb17.lib`, `acge17.lib`, `achapi17.lib`, `acismobj17.lib`, `adui17.lib`, `acui17.lib`.
4. Для конфигураций Debug/Release используйте Unicode-сборку (`Character Set = Use Unicode Character Set`) и убедитесь, что заданы `_UNICODE`, `UNICODE` и `AD_UNICODE`.
5. Строковые литералы для вызовов ObjectARX передавайте как `ACHAR` (например, через `ACRX_T("...")`) или через явную конвертацию ACP/UTF-8 → `ACHAR`.
6. Соберите `.arx`, загрузите через `APPLOAD` в AutoCAD 2007.
7. Выполните `SHOWDWGPROPS`.
8. Для импорта `XLSX2DWGPROP` откроется стандартный Windows-диалог выбора файла `.xlsx`.

## Примечание

- Код сделан без зависимостей от MFC/ATL-палитр (`CAdUiPaletteSet`), что уменьшает риск конфликтов линковки в старых toolchain.
- Модуль рассчитан на интеграцию в проект, где entrypoint уже создан через `IMPLEMENT_ARX_ENTRYPOINT(...)`.
- Команда `XLSX2DWGPROP` читает Excel-файл через COM; лист и колонки настраиваются в `Xlsx2DwgProp.ini` (`WorksheetIndex` или `WorksheetName`, `KeyColumn`, `ValueColumn`). Рекомендуется `WorksheetIndex`, чтобы избежать проблем кодировки.
- В `Xlsx2DwgProp.ini` можно выбрать программу чтения XLSX: `Reader=Excel` (через COM) или `Reader=LibreOffice` (через `soffice --headless --convert-to csv`), путь задается параметром `LibreOfficePath`.
- Для `Reader=LibreOffice` добавлен поиск исполняемого файла по типовым путям (`soffice.exe`, `C:\\Program Files\\LibreOffice\\program\\soffice.exe`, `C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe`) и вывод детального сообщения об ошибке в командной строке.
- Рекомендуемое значение `LibreOfficePath`: `C:\\Program Files\\LibreOffice\\program\\soffice.exe` (если указан `scalc.exe`, модуль автоматически попробует соседний `soffice.exe`).
- В верхней части панели есть строка поиска: фильтрация выполняется сразу по `Tag`, `Name` (описание из колонки C) и `Value`.
- При импорте из XLSX также читается колонка `C` как описание атрибута: в панели оно отображается справа от тега свойства (ключ из колонки `B`) более мелким шрифтом.
- При импорте также читается колонка `A` как название группы параметров. На форме есть выпадающий список групп (по умолчанию `Все группы`) для фильтрации отображаемых параметров по выбранной группе.
- Строка статуса XLSX размещена самой верхней строкой формы, чтобы не перекрываться фильтрами.
- Для длинных тегов в панели используется двухколоночная отрисовка (тег/описание) с `...`, чтобы текст не накладывался.
- После импорта XLSX полный путь и CRC32 файла сохраняются в пользовательских свойствах самого DWG (служебные ключи `__MG_XLSX_LAST_FILE_FULLPATH` и `__MG_XLSX_LAST_FILE_CRC32`), поэтому проверка хеша привязана к конкретному чертежу, а не к общему ini-файлу проекта. Период проверки (`XLSX.HashCheckMinutes`, по умолчанию 10 минут) по-прежнему читается из `DwgPropsPanel.ini`.
- Убедитесь, что `src/Xlsx2DwgProp.cpp` добавлен в проект/solution (иначе будет LNK2019 по `Xlsx2DwgProp_Command`).
- Убедитесь, что `src/ToolbarSetup.cpp` добавлен в проект/solution (иначе будет LNK2019 по `EnsureMgPanelToolbar`).
- Для toolbar используется MENUGROUP `MG-Project`: при первом запуске (если рядом с `.arx` нет ни `MG-Project.cuix`, ни `MG-Project.cui`, ни `MG-Project.mns`) модуль создаёт базовый `MG-Project.mns`, затем пытается загрузить `MG-Project` из `MG-Project.cuix` / `MG-Project.cui` / `MG-Project.mns`.
- После первого создания `MG-Panel` выполняется сохранение menugroup (`Save`), чтобы положение/состояние toolbar сохранялось между перезапусками AutoCAD.
- Для иконок кнопок toolbar ожидаются файлы рядом с `.arx`: `Loadpanel16.bmp`, `Loadpanel32.bmp`, `UnLoadpanel16.bmp`, `UnLoadpanel32.bmp` (иконки применяются и к уже существующей `MG-Panel`).
- При инициализации модуля вызывается `_tsetlocale(LC_ALL, _T("russian"))` для улучшения обработки кириллицы в старом toolchain.
- Сообщения `XLSX2DWGPROP` на русском выводятся через runtime-конвертацию UTF-8 -> ACP (`PrintUtf8`), т.к. одного `_tsetlocale(LC_ALL, _T("russian"))` в старом AutoCAD/VC++ часто недостаточно.
- Для toolbar используются только текстовые кнопки `SHOW/HIDE` (без генерации иконок).
- Положение и размер окна `DWG Properties` сохраняются в `%APPDATA%\\AutoCAD_Info_Panel\\DwgPropsPanel.ini` (при отсутствии файл создается со значениями по умолчанию) и восстанавливаются при следующем открытии панели.
- Из системных переменных в списке скрыты `DBMOD`, `LOGINNAME`, `DWGTITLED`.
