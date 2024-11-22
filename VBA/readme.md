# Materials for Excel VBA programming course
## in processing

## Application
```
### Application.ScreenUpdating property (Excel)
[microsoft](https://learn.microsoft.com/en-us/office/vba/api/excel.application.screenupdating)

EN: The code turns off / on screen updating. Makes your code run faster.
PL: Wstrzymanie odświeżania ekranu w Excel VBA. Przyspiesza działanie kodu.
```
```
Application.ScreenUpdating = False
Application.ScreenUpdating = True                                      
```                                      

```
Application.UseSystemSeparators = False
Application.DecimalSeparator = "."			
Application.UseSystemSeparators = True	
```
```
Application.DisplayAlerts = FALSE
Application.DisplayAlerts = TRUE
```
Application.Calculation = xlManual
Application.Calculation = xlAutomatic
```
```
Application.AutoRecover.Enabled = True
Application.AutoRecover.Time = 5 
```	
ActiveWorkbook.AutoSaveOn = False
ActiveWorkbook.AutoSaveOn = True
```
