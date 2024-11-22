# Materials for Excel VBA programming course

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
```
Application.UseSystemSeparators = False
Application.DecimalSeparator = "."			
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
```
Application.Calculation = xlManual
Application.Calculation = xlAutomatic
```
### Application.AutoRecover property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.application.autorecover)

```
Application.AutoRecover.Enabled = True
Application.AutoRecover.Time = 5 
```	
### Application.AutoSaveOn property
```
ActiveWorkbook.AutoSaveOn = False
ActiveWorkbook.AutoSaveOn = True
```
