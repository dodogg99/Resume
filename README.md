## 檔案說明
* ### 刀具使用分析:
  #### 將機台使用的刀具數據進行樞紐分析，從換刀原因、產品、刀具類型等分類來找異常情況。
* ### 刀具呆滯分析: 
  #### 將刀具庫存及領用紀錄進行分析，找出呆滯刀具及曾經加工的產品並統計總庫存金額、呆滯刀具金額及本月領用金額。
  - 引用VBAProject : Microsoft Scripting Runtime
* ### 即時生產機況:
  #### 自動從網頁下載機台生產資料，整理成各個產品在哪些機台加工、生產機台總數的即時訊息。
  - 引用VBAProject : Microsoft Scripting Runtime
* ### 存貨呆滯分析:
  #### 將兩個月份的存貨進行篩選比較，找出異常增加的存貨項目及原因。
  - 引用VBAProject : Microsoft Scripting Runtime
* ### 績效數據下載及整理:
  #### 使用Selenium爬取資料庫網頁API的績效數據，更新每月的績效指標。
  - 引用VBAProject : Selenium Type Library
  - Selenium安裝
  - chromedriver需與chrome相同版本

## 使用VBA語法
  - QueryTable、Calling Sub、Do Until Loop、For Loop、If Else Statement、 FormulaArray、Autofilter、PivotTable、FindElementByXpath

## 存貨呆滯分析範例
* ### 步驟說明:
  #### 1.打開存貨呆滯分析.xlsm檔案
  #### 2.修改變數名稱工作表中**B11-B18**欄位
  #### 3.執行巨集即可得到分析結果

* ### 變數名稱工作表
![變數名稱](https://github.com/dodogg99/VBA-project/blob/main/%E5%AD%98%E8%B2%A8%E5%91%86%E6%BB%AF%E5%88%86%E6%9E%90-%E8%AE%8A%E6%95%B8%E5%90%8D%E7%A8%B1%E8%A1%A8.JPG)
  
* ### 程式碼  
```vba
Option Explicit
Sub 不同月份存貨差異比較()
    ChDir Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B7").Value '設定存貨呆滯檔案資料夾位置
    Dim last_month_file As String, analysis_file As String, last_month As String, this_month As String, department _
     As String, warehouse_name As String, inventory_type As String, inventory_type_name As String, this_month_final_row As String, last_month_final_row As String
    Dim warehouse_address As Range, inventory_address As Range
    Dim department_detail_name As Dictionary
    Set department_detail_name = New Dictionary
    '讀取變數名稱裡的工作表名稱
    department_detail_name.Add Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B3").Value, Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B4").Value
    department_detail_name.Add Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B5").Value, Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B6").Value

    '輸入要分析的項目
    department = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B16").Value '要分析的部門，要跟部門工作表名稱相同
    last_month_file = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B13").Value '上個月的滯料分析表檔名
    analysis_file = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B15").Value '另存的差異分析檔名
    this_month = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B12").Value
    last_month = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B11").Value
    warehouse_name = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B17").Value '倉別名稱
    inventory_type = Workbooks("存貨呆滯分析").Worksheets("變數名稱").Range("B18").Value '存貨類別
    inventory_type_name = "類別代碼"

    '進行這個月存貨類別及倉別篩選
    Workbooks(analysis_file).Sheets(department_detail_name(department)).Activate
    Set warehouse_address = Workbooks(analysis_file).Sheets(department_detail_name(department)).Range("a1:z1").Find(warehouse_name, LookIn:=xlValues)
    Set inventory_address = Workbooks(analysis_file).Sheets(department_detail_name(department)).Range("a1:z1").Find(inventory_type_name, LookIn:=xlValues)
    this_month_final_row = Range("A1").CurrentRegion.Rows.Count
    Range("A1:Z1").AutoFilter Field:=inventory_address.Column, Criteria1:=inventory_type
    Range("A1:Z1").AutoFilter Field:=warehouse_address.Column, Criteria1:="<>0"
    If WorksheetFunction.Subtotal(3, Range(Cells(1, inventory_address.Column), Cells(this_month_final_row, inventory_address.Column))) > 1 Then
       '將這月份存貨項目篩選值貼上差異分析檔
        Sheets.Add After:=Sheets(department)
        Sheets(department).Next.Name = department & "整理"
        Worksheets(department_detail_name(department)).Range("A1").CurrentRegion.Copy Worksheets(department & "整理").Range("A1")
        Range("1:1").Delete
        Union(Range(Cells(1, 3), Cells(Rows.Count, warehouse_address.Column - 2)), Range(Cells(1, warehouse_address.Column + 1), Cells(Rows.Count, 26))).Delete
        Worksheets(department & "整理").Range("A1").CurrentRegion.Copy Worksheets(department).Range("m2")
        Workbooks(analysis_file).Sheets(department).Activate
        Range("m1").Value = ("料號")
        Range("n1").Value = ("倉別")
        Range("o1").Value = (this_month & "數量")
        Range("p1").Value = (this_month & "金額")
        Range("q1").Value = (last_month & "數量")
        Range("r1").Value = (last_month & "金額")
        Range("s1").Value = ("數量差值")
        Range("t1").Value = ("金額差值")
        Application.DisplayAlerts = False
        Sheets(department & "整理").Delete
        Application.DisplayAlerts = True
    End If
    '取消部門明細工作表篩選
    Workbooks(analysis_file).Sheets(department_detail_name(department)).Activate
    Range("1:1").AutoFilter
    
    '進行上個月存貨類別及倉別篩選
    Workbooks(last_month_file).Sheets(department_detail_name(department)).Activate
    Range("A1:Z1").AutoFilter Field:=inventory_address.Column, Criteria1:=inventory_type
    Range("A1:Z1").AutoFilter Field:=warehouse_address.Column, Criteria1:="<>0"
    last_month_final_row = Range("A1").CurrentRegion.Rows.Count
    If WorksheetFunction.Subtotal(3, Range(Cells(1, inventory_address.Column), Cells(last_month_final_row, inventory_address.Column))) > 1 Then
        '將上月份存貨項目篩選值貼上差異分析檔
        Sheets.Add After:=Sheets(department)
        Sheets(department).Next.Name = department & "整理"
        Sheets(department_detail_name(department)).Range("A1").CurrentRegion.Copy Worksheets(department & "整理").Range("A1")
        Range("a1").EntireRow.Delete
        Union(Range(Cells(1, 3), Cells(Rows.Count, warehouse_address.Column - 2)), Range(Cells(1, warehouse_address.Column + 1), Cells(Rows.Count, 26))).Delete
        Worksheets(department & "整理").Range("A1").CurrentRegion.Copy Workbooks(analysis_file).Worksheets(department).Range("u2")
        Application.DisplayAlerts = False
        Sheets(department & "整理").Delete
        Application.DisplayAlerts = True
        Workbooks(analysis_file).Worksheets(department).Activate
        Range("U1").Value = ("料號")
        Range("v1").Value = ("倉別")
        Range("w1").Value = (last_month & "數量")
        Range("x1").Value = (last_month & "金額")
        Range("y1").Value = (this_month & "數量")
        Range("z1").Value = (this_month & "金額")
        Range("aa1").Value = ("數量差值")
        Range("ab1").Value = ("金額差值")
    End If
    Workbooks(last_month_file).Sheets(department_detail_name(department)).Activate
    Range("1:1").AutoFilter

    '比較兩個月的差值
    Dim this_month_data_row As Long, last_month_data_row As Long
    Workbooks(analysis_file).Sheets(department).Activate
    If Not Range("m2") = "" And Not Range("u2") = "" Then
        this_month_data_row = Range("M1").End(xlDown).Row
        last_month_data_row = Range("u1").End(xlDown).Row
        Range("q2").FormulaArray = "=VLOOKUP(RC[-4]&RC[-3],IF({1,0},R2C[4]:R" & last_month_data_row & "C[4]&R2C[5]:R" & last_month_data_row & "C[5],R2C[6]:R" & last_month_data_row & "C[6]),2,0)"
        Range("r2").FormulaArray = "=VLOOKUP(RC[-5]&RC[-4],IF({1,0},R2C[3]:R" & last_month_data_row & "C[3]&R2C[4]:R" & last_month_data_row & "C[4],R2C[6]:R" & last_month_data_row & "C[6]),2,0)"
        Range("s2").FormulaR1C1 = "=RC[-4]-RC[-2]"
        Range("t2").FormulaR1C1 = "=RC[-4]-RC[-2]"
        Range("q2:t2").AutoFill Destination:=Range("q2:t" & this_month_data_row)
        Range("M1:T" & this_month_data_row).Borders.LineStyle = xlContinuous
        Workbooks(analysis_file).Worksheets(department).Range("M1").ColumnWidth = 16
        Range("y2").FormulaArray = "=VLOOKUP(RC[-4]&RC[-3],IF({1,0},R2C[-12]:R" & this_month_data_row & "C[-12]&R2C[-11]:R" & this_month_data_row & "C[-11],R2C[-10]:R" & this_month_data_row & "C[-10]),2,0)"
        Range("z2").FormulaArray = "=VLOOKUP(RC[-5]&RC[-4],IF({1,0},R2C[-13]:R" & this_month_data_row & "C[-13]&R2C[-12]:R" & this_month_data_row & "C[-12],R2C[-10]:R" & this_month_data_row & "C[-10]),2,0)"
        Range("aa2").FormulaR1C1 = "=RC[-4]-RC[-2]"
        Range("ab2").FormulaR1C1 = "=RC[-4]-RC[-2]"
        Range("y2:ab2").AutoFill Destination:=Range("y2:ab" & last_month_data_row)
        Range("u1:ab" & last_month_data_row).Borders.LineStyle = xlContinuous
        Workbooks(analysis_file).Worksheets(department).Range("U1").ColumnWidth = 16
    '找不到篩選項目時，告知是哪個月份
    ElseIf Range("m2") = "" And Range("u2") = "" Then
        MsgBox this_month & last_month & "沒有篩選值無法比較"
    ElseIf Range("m2") = "" Then
        MsgBox this_month & "沒有篩選值無法比較"
    Else
        MsgBox last_month & "沒有篩選值無法比較"
    End If
End Sub
```
