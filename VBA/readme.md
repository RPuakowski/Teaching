# Materials for Excel VBA programming course

-------------------------------------------------------------------------------------------------------------------

## Application object (Excel)
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application(object))

EN: Represents the entire Microsoft Excel application.

PL: Reprezentuje całą aplikację Microsoft Excel.

### Application.ScreenUpdating property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application.screenupdating)

EN: The code turns off / on screen updating. Makes your code run faster.

PL: Wstrzymanie odświeżania ekranu w Excel VBA. Przyspiesza działanie kodu.

```
Application.ScreenUpdating = False
Application.ScreenUpdating = True                                      
```                                      

### Application.UseSystemSeparators property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usesystemseparators)

EN: Allows to change application default separators.

PL: Pozwala zmienić domyślne separator aplikacji.

```
Application.UseSystemSeparators = False
Application.DecimalSeparator = "."
Application.ThousandsSeparator = "-" 			
Application.UseSystemSeparators = True
```

### Application.DisplayAlerts property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts)

EN: True if Microsoft Excel displays certain alerts and messages while a macro is running.

Note: does not work from immediate window.

PL: PRAWDA jeżeli Microsoft Excel wyśtwiela pewne ostrzeżenia i powiadomienia w trakcie pracy makra.

Zauważ: nie działa z poziomu okna bezpośredniego (immediate window).
    
```
Application.DisplayAlerts = FALSE
Application.DisplayAlerts = TRUE
```

### Application.Calculation property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)

EN: Allows to switch calculation mode from automatic to manual and vice versa.

Note: Useful when working with huge files which slows your machine down.

PL: Pozwala zameinić tryb przeliczania danych z automatycznego na ręczny i na odwrót.

Zauważ: Przydatne kiedy pracujesz z dużymi plikami, które zwalniają pracę urządzenia.

```
Application.Calculation = xlManual
Application.Calculation = xlAutomatic
```

### Application.AutoRecover property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application.autorecover)

EN: Allows to manipulate properties of autosave mechanism. 

PL: Pozwala manipulować właściwościami mechanizmu autozapisu

```
Application.AutoRecover.Enabled = True
Application.AutoRecover.Time = 5
Application.AutoRecover.Path = "C:\"  
```

-------------------------------------------------------------------------------------------------------------------

## Workbook object
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook)

EN: Represents a Microsoft Excel workbook.

PL: Reprezentuje dany skoroszyt aplikacji.

```
? Workbooks.Count
Workbooks(1).Activate
Workbooks("Book1.xlsx").Close
ActiveWorkbook.SaveAs Filename:= "MYFILENAME.xlsx"
```

### Workbook.AutoSaveOn property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autosaveon)

EN: Allows user to automatically save given workbook.

PL: Pozwala użytkownikowi automatycznie zapisywać dany skoroszyt.

```
ActiveWorkbook.AutoSaveOn = False
ActiveWorkbook.AutoSaveOn = True
```

TBA
