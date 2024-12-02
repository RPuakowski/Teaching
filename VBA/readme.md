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
I recommend you read this tutorial by Paul Kelly.
[link](https://trumpexcel.com/vba-workbook/)

### Workbook.AutoSaveOn property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autosaveon)

EN: Allows user to automatically save given workbook.

PL: Pozwala użytkownikowi automatycznie zapisywać dany skoroszyt.

```
ActiveWorkbook.AutoSaveOn = False
ActiveWorkbook.AutoSaveOn = True
```

### Workbook.Protect / Unprotect method
                                    
ThisWorkbook.Protect
ThisWorkbook.Unprotect 

workbook.Protect([Password], [Structure], [Windows])
 
```
ThisWorkbook.Protect Password:="password"
ThisWorkbook.Unprotect Password:="password"

ThisWorkbook is the workbook where the running code is stored. 
	
ActiveWorkbook.Unprotect
ActiveWorkbook.Protect Password:="password"
ActiveWorkbook.Unprotect Password:="password"
	
ActiveWorkbook.Protect Password:=InputBox("Enter a protection password:")	
ActiveWorkbook.UnProtect Password:=InputBox("Enter a protection password:")	
ActiveWorkbook.Protect "password"
```
	
TBA

Workbook.Activate

Using Workbook Index numbers

Workbook.Close

ActiveWorkbook

ThisWorkbook

Workbooks.Add

Workbooks.Open

Workbooks.Save

-------------------------------------------------------------------------------------------------------------------

## Worksheet object
[link](https://learn.microsoft.com/en-us/office/vba/api/sheet)

EN: Represents a sheet in a workbook.

PL: Reprezentuje dany arkusz w skoroszycie.

Remember! / Pamiętaj!

    Sheets = Worksheets + Charts


### Sheets.Visible property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.visible)

EN: Returns or sets an XlSheetVisibility value that determines whether the object is visible.

PL: Zwraca bądź ustawia własnośc XlSheetVisibility, która wskazuje czy arkusz jest widoczny.

```
Sheets("Sheet1").Visible = -1
Sheets("Sheet1").Visible =  0
Sheets("Sheet1").Visible =  2    
	
Sheets("Sheet1").Visible = xlSheetVisible 
Sheets("Sheet1").Visible = xlSheetHidden
Sheets("Sheet1").Visible = xlSheetVeryHidden

Sheets("Sheet1").Visible = FALSE
Sheets("Sheet1").Visible = TRUE
```

### Sheets.Index property
[link](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.index)
[link](https://learn.microsoft.com/en-us/office/vba/excel/concepts/workbooks-and-worksheets/refer-to-sheets-by-index-number)

An index number is a sequential number assigned to a sheet, based on the position of its sheet tab (counting from the left) among sheets of the same type. The following procedure uses the Worksheets property to activate the first worksheet in the active workbook.
		
Each sheet has its own sequential number (index) in the collection of all sheets. You can use the index to construct worksheet references.
		

``` 
? TypeName(ActiveSheet)
		
? Sheets.Count
? Worksheets.count
? charts.count
		
? Sheet1.Index
? Sheets("Sheet1").Index
		
Sheets("Sheet1").Index = 4 ' invalid property assingment
``` 
		
The index of sheets is the left-to-right order in which they appear.

You can change the index by reordering the sheet.

```
Sheets("Sheet1").Move Before:=Sheets(1)
Sheets("Sheet1").Move Before:=Sheets(2)
Sheets("Sheet1").Move Before:=Sheets(3)
		
Sheets("Sheet1").Move After:=Sheets(3)
Sheets("Sheet1").Move After:=Sheets(2)	
Sheets("Sheet1").Move After:=Sheets(1)		 
```

### Sheet.Select method

```
Worksheets(2).Select
Sheets(2).Select
Sheet2.Select
ActiveSheet.Select
Sheets.Select
```
	
### Sheets.Add method

 ```
Sheets.Add 
Sheets.Add Before:=Sheets(1)
Sheets.Add After:=Sheets(2)
Sheets.Add After:=Sheets(Sheets.Count)
Sheets.Add.Name = "IlikeCats" 
Sheets.Add(After:=Sheets("IlikeCats")).Name = "IlikeDogs"
```
	
More: https://www.automateexcel.com/vba/add-and-name-worksheets/
	
### Sheet.Delete method

 ```
Sheet1.Delete
Sheet(5).Delete
Sheets("IlikeCats").Delete
Sheets("IlikeDogs").Delete
ActiveSheet.Delete
```

### Sheet.Name property
	
```
Sheet13.Name = "IlikeCats"
Sheets("Sheet9").Name = "Sheet12"
Sheets(2).Name = "IlikeDogs"				
Sheets("IlikeCats").Name = Sheets("IlikeCats").Range("A1")
```

### Sheet.Copy method
	
```
Sheets("IlikeMice").Copy
Sheets("IlikeMice").Copy After:=Sheets(1)
Sheets("IlikeMice").Copy Before:=Sheets("IlikeDogs")
Sheets("IlikeMice").Copy After:=Sheets(Sheets.Count)
Sheets("IlikeMice").Copy 
ActiveSheet.Name = "Sheet1"		
```
	
### Sheet.Protect method
	
```
Sheets("Sheet1").Protect
Sheets("Sheet1").Protect Password:="myPassword"
Sheets("Sheet1").Protect Password:=InputBox("Enter a protection password:")	
```
	
### Sheet.UnProtect method
	
```
Sheets("Sheet1").UnProtect
Sheets("Sheet1").UnProtect Password:="myPassword"
Sheets("Sheet1").UnProtect Password:=InputBox("Enter a protection password:")	
```

-------------------------------------------------------------------------------------------------------------------

