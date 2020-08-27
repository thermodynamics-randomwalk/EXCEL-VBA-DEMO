# vba content

1. 将一张大表分拆为小表

   ```vb
   Option Explicit
   Sub ShtAdd2()
       MsgBox "下面将根据G列的分行名新建不同的工作表。"
       Dim i As Integer, sht As Worksheet
       i = 15                                  '第一条记录的行号为15
       Set sht = Worksheets("汇总")
       Do While sht.Cells(i, "G") <> ""     '定义循环条件
           On Error Resume Next
           If Worksheets(sht.Cells(i, "G").Value) Is Nothing Then
           Worksheets.Add after:=Worksheets(Worksheets.Count)      '在所有工作表后插入新工作表
           ActiveSheet.Name = sht.Cells(i, "G").Value '更改工作表的标签名称
           End If
           i = i + 1                         '行号增加1
       Loop
   End Sub
   
   Sub split_sheets()
       MsgBox "下面将内容分按分行分到各个工作表中!"
       Dim i As Long, bj As String, rng As Range
       i = 15
       bj = Cells(i, "G").Value
       Do While bj <> ""
           '将分表中A列第一个空单元格赋给rng
           Set rng = Worksheets(bj).Range("A65536").End(xlUp).Offset(1, 0)
           Cells(i, "A").Resize(1, 23).Copy rng   '将记录复制到相应的工作表中
           i = i + 1
           bj = Cells(i, "G").Value
       Loop
   End Sub
   ```

2. 将大表标题行赋予小表

   ```vb
   Sub title_distribute()
   Dim sh As Worksheet, ir As Integer
   For Each sh In Worksheets
       If sh.Name <> "不良明细" And sh.Name <> Sheet1.Name Then
       ir = sh.Range("A65536").End(xlUp).Row
       Sheets("汇总").Rows("1:14").Copy sh.Range("a1")
       Sheets("汇总").Rows("4581").Copy sh.Range("a" & ir + 1)
       End If
   Next
   End Sub
   ```

3. 将分拆的小表拆成独立文件形式

   ```vb
   Sub SaveToFile()
       MsgBox "下面将把各个工作表保存为单独的工作薄文件，" & Chr(13) _
            & "保存在当前文件夹下的“分行各个表格”文件夹下！"
       Application.ScreenUpdating = False                        '关闭屏幕更新
       Dim folder As String
       folder = ThisWorkbook.Path & "\各个分行"
       If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder    '如果文件夹不存在，新建文件夹
       Dim sht As Worksheet
       For Each sht In Worksheets                                 '遍历工作表
           sht.Copy                                               '复制工作表到新工作薄
           ActiveWorkbook.SaveAs folder & "\" & sht.Name & ".xls" '保存工作薄到指定文件夹，并命名
           ActiveWorkbook.Close
       Next
       Application.ScreenUpdating = True                          '开启屏幕更新
   End Sub
   ```

4. 将独立文件进行汇总

   ```vb
   Sub aggregate_Wb()
       Dim bt As Range, r As Long, c As Long
       r = 14    '1 是表头的行数
       c = 23    '8 是表头的列数
       Range(Cells(r + 1, "A"), Cells(65536, c)).ClearContents    ' 清除汇总表中原表数据
       Application.ScreenUpdating = False
       Dim FileName As String, wb As Workbook, Erow As Long, fn As String, arr As Variant, SHT As Worksheet
       FileName = Dir(ThisWorkbook.Path & "\*.xls")
       Do While FileName <> ""
           If FileName <> ThisWorkbook.Name Then    ' 判断文件是否是本工作簿
               Erow = Range("A14").CurrentRegion.Rows.Count + 1    ' 取得汇总表中第一条空行行号
               fn = ThisWorkbook.Path & "\" & FileName
               Set wb = GetObject(fn)    ' 将fn 代表的工作簿对象赋给变量
               Set SHT = wb.Worksheets(1)    ' 汇总的是第1 张工作表
               ' 将数据表中的记录保存在arr 数组里
               arr = SHT.Range(SHT.Cells(r + 1, "A"), SHT.Cells(65536, "B").End(xlUp).Offset(0, 23))
               ' 将数组arr 中的数据写入工作表
               Cells(Erow, "A").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
               wb.Close False
           End If
           FileName = Dir    ' 用Dir 函数取得其他文件名，并赋给变量
       Loop
       Application.ScreenUpdating = True
   End Sub
   ```

5. 清除每个表中内容

   ```vb
   Sub clear_contents()
   	Dim sh As Worksheet
   	For Each sh In Worksheets
   			sh.Rows("15:65536").clear
   	Next
   End Sub
   ```

6. 清除每个表中相同的行

   ```vb
   Sub qc()
   Dim sh As Worksheet, ir As Integer
   For Each sh In Worksheets
       ir = sh.Range("A65536").End(xlUp).Row
       sh.Rows("1:3").Delete
       sh.Rows(ir + 1).Delete
       sh.Range("A65536").End(xlUp).Offset(-6, 0).Resize(7, 256).Delete '倒数7行，工作簿不能存在空白工作表
   Next
   End Sub
   
   '清除每个表中冻结窗口
   
   Sub wc()
   Dim sh As Worksheet, i As Integer
   For i = 1 To Sheets.Count
      Sheets(i).Select
      Range("a1").Select
      ActiveWindow.FreezePanes = False
   Next i
   End Sub
   ```

7. 合并工具（将单个文件转到一张表格中)

   ```vb
   ' XLS => XLSX
   Option Explicit
   Sub sConvertXLS()
       Dim sPath As String
       Dim sName As String
       Dim sFile As String
       Dim sExt As String
       Dim sNewExt As String
       Dim sFirstFile As String
       Dim objWorkbook As Workbook
       Dim iFormat As Integer
       Dim lFileCount As Long
       With Application.FileDialog(msoFileDialogFolderPicker)
           .Show
           If .SelectedItems.Count = 0 Then
               MsgBox "请选择文件目录!", vbInformation, "Excel Home"
               Exit Sub
           End If
           sPath = .SelectedItems(1) & "\"
       End With
       sName = Dir(sPath & "*.xl*")
       If Len(sName) > 0 Then
           Application.DisplayAlerts = False
           Application.ScreenUpdating = False
           sFirstFile = sName
           Do
               sExt = UCase(Right(sName, 4))
               Select Case sExt
               Case ".XLS", ".XLA", ".XLT"
                   Set objWorkbook = Workbooks.Open(sPath & sName)
                   sFile = Left(sName, Len(sName) - 4)
                   Select Case sExt
                   Case ".XLS"
                       If fWorkbookWithCode(objWorkbook) Then
                           sNewExt = ".xlsm"
                           iFormat = xlOpenXMLWorkbookMacroEnabled
                       Else
                           sNewExt = ".xlsx"
                           iFormat = xlOpenXMLWorkbook
                       End If
                   Case ".XLA"
                       sNewExt = ".xlam"
                       iFormat = xlOpenXMLAddIn
                   Case ".XLT"
                       If fWorkbookWithCode(objWorkbook) Then
                           sNewExt = ".xltm"
                           iFormat = xlOpenXMLTemplateMacroEnabled
                       Else
                           sNewExt = ".xltx"
                           iFormat = xlOpenXMLTemplate
                       End If
                   End Select
                   If Not objWorkbook Is Nothing Then
                       With objWorkbook
                           .SaveAs sPath & sFile & sNewExt, iFormat
                           .Close
                           lFileCount = lFileCount + 1
                       End With
                   End If
                   Set objWorkbook = Nothing
               End Select
               sName = Dir
           Loop While Len(sName) > 0 And sName <> sFirstFile
           Application.DisplayAlerts = True
           Application.ScreenUpdating = True
           If lFileCount > 0 Then
               MsgBox "成功转换 " & lFileCount & " 个文件!", _
                      vbInformation, "Excel Home"
           Else
               MsgBox "没有需要转换的文件!", vbInformation, "Excel Home"
           End If
       Else
           MsgBox "没有Excel文件!", vbInformation, "Excel Home"
       End If
       Set objWorkbook = Nothing
   End Sub
   Function fWorkbookWithCode(objWb As Workbook) As Boolean
       Dim objVBC As Object
       Dim lCodeLines As Long
       For Each objVBC In objWb.VBProject.VBComponents
           lCodeLines = lCodeLines + _
                        objVBC.CodeModule.CountOfLines
       Next
       fWorkbookWithCode = (lCodeLines > 0)
   End Function
   ```

8. 为工作表建立目录

   ```vb
   Sub create_sht_list()
       MsgBox "下面将为工作薄中所有工作表建立目录!"
       Rows("2:65536").ClearContents                    '清除工作表中原有数据
       Dim sht As Worksheet, irow As Integer
       irow = 2                                         '在第2行写入第一条记录
       For Each sht In Worksheets                       '遍历工作表
           Cells(irow, "A").Value = irow - 1            '写入序号
           '写入工作表名，并建立超链接
           ActiveSheet.Hyperlinks.Add Anchor:=Cells(irow, "B"), Address:="", _
                SubAddress:="'" & sht.Name & "'!A1", TextToDisplay:=sht.Name
           irow = irow + 1                              '行号加1
       Next
   End Sub
   ```

9. 依照目录对工作表批量更改名称

   ```vb
   Sub RENAME()
       Dim SHTNAME as string, SHT As Worksheet, i As Integer
       On Error Resume Next
       For i = 2 To Cells(Rows.Count, "a").End(xlUp).Row
           SHTNAME = Cells(i, "b").Value
           Worksheets(SHTNAME).Name = Cells(i, "j").Value
       Next
   End Sub
   ```

10. 依据特定顺序排列工作表

    ```vb
    Sub get_sht_name_list() '将表格名称提取到a列
    Dim sht As Worksheet, k&
    Range("a:a").ClearContents
    Range("a:a").NumberFormat = "@"
    Cells(1, 1) = "list"
    k = 1
    For Each sht In Worksheets
        k = k + 1
        Cells(k, 1) = sht.Name
    Next
    End Sub
    
    Sub sortsheet()
    Dim sht As Worksheet, shtname as string, i&
    Set sht = ActiveSheet
    For i = 2 To sht.Cells(Rows.Count, 1).End(xlUp).Row
        shtname = sht.Cells(i, 1).Value
        Worksheets(shtname).Move after:=Worksheets(i - 1)
    Next
    sht.Activate
    End Sub
    ```

11. 显示文件夹的文件名称（存在dos形式 dir*.*/b>list.txt,如果要文件夹名称选用f = Dir(p & "各个*", vbDirectory)）

    ```vb
    Sub filedir()
    Dim p$, f$, k&
     With Application.FileDialog(msoFileDialogFolderPicker)
                     .AllowMultiSelect = False
                    If .Show Then
                    p = .SelectedItems(1)
                    Else
                    Exit Sub
                    End If
    End With
                    If Right(p, 1) <> "\" Then p = p & "\"
                    f = Dir(p & "*.*")
                    Range("a:a").ClearContents
                    Range("a1") = "list"
                    k = 1
                        Do While f <> ""
                        k = k + 1
                        Cells(k, 1) = f
                        f = Dir
                       Loop
    MsgBox "ok"
                           
    End Sub
    ```

12. 按指定条件批量删除excel工作簿

    ```vb
    Sub get_books_path()
    Dim p$, f$, k&
    With Application.FileDialog(msoFileDialogFolderPicker)
                    .AllowMultiSelect = False
                    If .Show Then p = .SelectedItems(1) Else: Exit Sub
    End With
    If Right(p, 1) <> "\" Then p = p & "\"
    [a:b].ClearContents
    k = 1
    [a1] = "list"
    [b1] = "delete?"
    f = Dir(p & "*.xls*")
    Do While f <> ""
        k = k + 1
        Cells(k, 1) = p & f
        f = Dir
    Loop
    End Sub
        
        Sub delete_books()
     Dim r, i&
     r = [a1].CurrentRegion
     For i = 2 To UBound(r)
        If r(i, 2) = "delete" Then Kill r(i, 1)
     Next
    End Sub
    ```

13. 按指定名称批量创建excel工作簿

    ```vb
    Sub createwks()
    Dim i&, p$, r
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    p = ThisWorkbook.Path & "\"
    r = [b1].CurrentRegion
    For i = 2 To UBound(r)
        With Workbooks.Add
            .SaveAs p & r(i, 1), xlWorkbookDefault
            .Close True
        End With
    Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    End Sub
    ```

14. 按指定条件汇总各分表数据到总表

    ```vb
    Sub CollectSheets()
        Dim sht As Worksheet, rng As Range, k&, trow&,temp
        Application.ScreenUpdating = False
        '取消屏幕更新，加快代码运行速度
        temp = InputBox("请输入需要合并的工作表所包含的关键词：", "提醒")
        If StrPtr(temp) = 0 Then Exit Sub
        '如果点击了inputbox的取消或者关闭按钮，则退出程序
        trow = Val(InputBox("请输入标题的行数", "提醒"))
        If trow < 0 Then MsgBox "标题行数不能为负数。", 64, "警告": Exit Sub
        '取得用户输入的标题行数，如果为负数，退出程序
        Cells.ClearContents
        '清空当前表数据
        For Each sht In Worksheets
        '循环读取表格
            If sht.Name <> ActiveSheet.Name Then
            '如果表格名称不等于当前表名则……
                If InStr(1, sht.Name, temp, vbTextCompare) Then
               '如果表中包含关键词则进行汇总动作(不区分关键词字母大小写）
                    Set rng = sht.UsedRange
                    '定义rng为表格已用区域
                    k = k + 1
                    '累计K值
                    If k = 1 Then
                    '如果是首个表格，则K为1，则把标题行一起复制到汇总表
                        rng.Copy
                        [a1].PasteSpecial Paste:=xlPasteValues
                    Else
                        '否则，扣除标题行后再复制黏贴到总表，只黏贴数值
                        rng.Offset(trow).Copy
                        Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1).PasteSpecial Paste:=xlPasteValues
                    End If
                End If
            End If    
          Next
        [a1].Activate
        '激活A1单元格
        Application.ScreenUpdating = True
        '恢复屏幕刷新
    
    End Sub
    ```

15. 汇总多个工作簿每个工作表名称包含指定关键词的数据到总表

    ```vb
    Sub Collectwks()
        'ExcelHome VBA编程学习与实践，看见星光
        Dim Sht As Worksheet, rng As Range, Sh As Worksheet
        Dim Trow&, k&, arr, brr, i&, j&, book&, a&
        Dim p$, f$, Headr, Keystr
        '
        With Application.FileDialog(msoFileDialogFolderPicker)
        '取得用户选择的文件夹路径
            .AllowMultiSelect = False
            If .Show Then p = .SelectedItems(1) Else Exit Sub
        End With
        If Right(p, 1) <> "\" Then p = p & "\"
        '
        Keystr = InputBox("请输入需要合并的工作表所包含的关键词：", "提醒")
        If StrPtr(Keystr) = 0 Then Exit Sub
        '如果点击了inputbox的取消或者关闭按钮，则退出程序
        Trow = Val(InputBox("请输入标题的行数", "提醒"))
        If Trow < 0 Then MsgBox "标题行数不能为负数。", 64, "警告": Exit Sub
        Set Sht = ActiveSheet
        Application.ScreenUpdating = False '关闭屏幕更新
        Cells.ClearContents
        Cells.NumberFormat = "@"
        '清空当前表数据并设置为文本格式
        ReDim brr(1 To 200000, 1 To 2)
        '定义装汇总结果的数组brr，最大行数为20万行，2列是临时的
        '
        f = Dir(p & "*.xls*") '开始遍历工作簿
        Do While f <> ""
            On Error Resume Next
            If f <> ThisWorkbook.Name Then '避免同名文件重复打开出错
                With GetObject(p & f)
                '以'只读'形式读取文件时，使?胓etobject方法会比workbooks.open稍快
                    For Each Sh In .Worksheets '遍历表
                        If InStr(1, Sh.Name, Keystr, vbTextCompare) Then
                        '如果表中包含关键词则进行汇总(不区分关键词字母大小写）
                            Set rng = Sh.UsedRange  '单元格为区域 range("a2:b3")
                            If rng.Count > 1 Then
                            '如果rng的单元格数量大于1……
                                book = book + 1 '标记一下是否首个Sheet,如果首个sheet，BOOK=1
                                a = IIf(book = 1, 1, Trow + 1) '遍历读取arr数组时是否扣掉标题行
                                arr = rng.Value '数据区域读入数组arr
                                If UBound(arr, 2) + 2 > UBound(brr, 2) Then
                                '动态调整结果数组brr的最大列数，避免明细表列数不一的情况。
                                    ReDim Preserve brr(1 To 200000, 1 To UBound(arr, 2) + 2)
                                End If
                                For i = a To UBound(arr) '遍历行
                                    k = k + 1 '累加记录条数
                                    brr(k, 1) = f '数组第一列放工作簿名称
                                    brr(k, 2) = Sh.Name '数组第二列放工作表名称
                                    For j = 1 To UBound(arr, 2) '遍历列
                                        brr(k, j + 2) = arr(i, j)
                                    Next
                                Next
                            End If
                        End If
                    Next
                    .Close False '关闭工作簿
                End With
            End If
            f = Dir '下一个表格
        Loop
        If k > 0 Then
            Sht.Select
            [a1].Offset(IIf(Trow = 0, 1, 0)).Resize(k, UBound(brr, 2)) = brr '放数据区域
            [a1].Resize(1, 2) = [{"来源工作簿名称","来源工作表名"}]
            MsgBox "汇总完成。"
        End If
        Application.ScreenUpdating = True '恢复屏幕更新
    End Sub
    
    
    '修改版（针对于单个单元格）
            
    Sub Collectwks()
        Dim sht As Worksheet, rng As Range, sh As Worksheet
        Dim Trow&, k&, arr, brr, crr, i&, j&, book&, a&
        Dim p$, f$, Headr, Keystr, Keystr1
        '
        With Application.FileDialog(msoFileDialogFolderPicker)
        '取得用户选择的文件夹路径
            .AllowMultiSelect = False
            If .Show Then p = .SelectedItems(1) Else Exit Sub
        End With
        If Right(p, 1) <> "\" Then p = p & "\"
        '
        Keystr = InputBox("请输入需要合并的工作表所包含的关键词：", "提醒")
        If StrPtr(Keystr) = 0 Then Exit Sub
        Keystr1 = InputBox("请输入需要合并的工作表所包含的非合并单元格（e.g:a1)：", "提醒")
        If StrPtr(Keystr1) = 0 Then Exit Sub
        '如果点击了inputbox的取消或者关闭按钮，则退出程序
        Set sht = ActiveSheet
        Application.ScreenUpdating = False '关闭屏幕更新
        Cells.ClearContents
        Cells.NumberFormat = "@"
        '清空当前表数据并设置为文本格式
    
        ReDim brr(1 To 200000, 1 To 30)
    
        '定义装汇总结果的数组brr，最大行数为20万行，2列是临时的
        '
        f = Dir(p & "*.xls*") '开始遍历工作簿
        Do While f <> ""
            On Error Resume Next
            If f <> ThisWorkbook.Name Then '避免同名文件重复打开出错
                With GetObject(p & f)
                '以'只读'形式读取文件时，使?胓etobject方法会比workbooks.open稍快
                    For Each sh In .Worksheets '遍历表
                        If InStr(1, sh.Name, Keystr, vbTextCompare) Then
                        '如果表中包含关键词则进行汇总(不区分关键词字母大小写）
                            Set rng = sh.Range(Keystr1)
                           
                                arr = rng.Value '数据区域读入数组arr
                                
                            
                                For i = 1 To UBound(arr) '遍历行
                                    k = k + 1 '累加记录条数
                                    brr(k, 1) = f '数组第一列放工作簿名称
                                    brr(k, 2) = sh.Name '数组第二列放工作表名称
                                    brr(k, 3) = arr'所选单元格必须为单个 不能为区域
                                    Next
                              
                        End If
                    Next
                    .Close False '关闭工作簿
                End With
            End If
            f = Dir '下一个表格
        Loop
        MsgBox k
           If k > 0 Then
            sht.Select
            [a1].Offset(1, 0).Resize(k, UBound(brr, 2)) = brr  '放数据区域
            [a1].Resize(1, 2) = [{"来源工作簿名称","来源工作表名"}]
            MsgBox "汇总完成。"
        End If
        Application.ScreenUpdating = True '恢复屏幕更新
    End Sub
    ```

16. 复制指定文件夹下多工作簿的工作表到汇总工作簿

    ```vb
    Sub CltSheets()
     
        Dim P$, Bookn$, Book$, Keystr1, Keystr2, Shtname$, K&
        Dim Sht As Worksheet, Sh As Worksheet
        On Error Resume Next
        With Application.FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            If .Show Then P = .SelectedItems(1) Else: Exit Sub
        End With
        If Right(P, 1) <> "\" Then P = P & "\"
        Keystr1 = InputBox("请输入工作簿名称所包含的关键词。" & vbCr & "关键词可以为空，如为空，则默认选择全部工作簿")
        If StrPtr(Keystr1) = 0 Then Exit Sub '如果用户点击了取消或关闭按钮，则退出程序
        Keystr2 = InputBox("请输入工作表名称所包含的关键词。" & vbCr & "关键词可以为空，如为空，则默认选择符合条件工作簿的全部工作表")
        If StrPtr(Keystr2) = 0 Then Exit Sub
        Set Sh = ActiveSheet '当前工作表，赋值变量,代码运行完毕后，回到此表
        Bookn = Dir(P & "*.xls*")
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Do While Bookn <> ""
            If Bookn = ThisWorkbook.Name Then
                MsgBox "注意：指定文件夹中存在和当前表格重名的工作簿！！" & vbCr & "该工作簿无法打开，工作表无法复制。"
                '当出现重名工作簿时，提醒用户。
            Else
                If InStr(1, Bookn, Keystr1, vbTextCompare) Then
                '工作簿名称是否包含关键词，关键词不区分大小写
                    With GetObject(P & Bookn)
                        For Each Sht In .Worksheets
                            If InStr(1, Sht.Name, Keystr2, vbTextCompare) Then
                            '工作表名称是否包含关键词，关键词不区分大小写
                                If Application.CountIf(Sht.UsedRange, "<>") Then
                                '如果表格存在数据区域
                                    Shtname = Split(Bookn, ".xls")(0) & "-" & Sht.Name
                                    '复制来的工作表以"工作簿-工作表"形式起名。
                                    ThisWorkbook.Sheets(Shtname).Delete
                                    '如果已存在相关表名，则删除
                                    Sht.Copy after:=ThisWorkbook.Worksheets(Sheets.Count)
                                    K = K + 1
                                    '复制Sht到代码所在工作簿所有工作表的后面，并累计个数
                                    ActiveSheet.Name = Shtname
                                    '工作表命名。
                                End If
                            End If
                        Next
                        .Close False '关闭工作簿
                    End With
                End If
            End If
            Bookn = Dir '下一个符合条件的文件
        Loop
        Sh.Select '回到初始工作表
        MsgBox "工作表收集完毕，共收集：" & K & "个"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Sub
    
    '复制指定文件夹下多工作簿的工作表到汇总工作簿(依据指定单元格区域内是否有关键字）
                
    Sub CltSheets()
        
        Dim P$, Bookn$, Book$, Keystr1, Keystr2, Shtname$, K&
        Dim Sht As Worksheet, Sh As Worksheet
        On Error Resume Next
        With Application.FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            If .Show Then P = .SelectedItems(1) Else: Exit Sub
        End With
        If Right(P, 1) <> "\" Then P = P & "\"
        Keystr1 = InputBox("请输入工作簿名称所包含的关键词。" & vbCr & "关键词可以为空，如为空，则默认选择全部工作簿")
        If StrPtr(Keystr1) = 0 Then Exit Sub '如果用户点击了取消或关闭按钮，则退出程序
        Keystr2 = InputBox("请输入工作表名称所包含的关键词。" & vbCr & "关键词可以为空，如为空，则默认选择符合条件工作簿的全部工作表")
        If StrPtr(Keystr2) = 0 Then Exit Sub
        Set Sh = ActiveSheet '当前工作表，赋值变量,代码运行完毕后，回到此表
        Bookn = Dir(P & "*.xls*")
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Do While Bookn <> ""
            If Bookn = ThisWorkbook.Name Then
                MsgBox "注意：指定文件夹中存在和当前表格重名的工作簿！！" & vbCr & "该工作簿无法打开，工作表无法复制。"
                '当出现重名工作簿时，提醒用户。
            Else
                If InStr(1, Bookn, Keystr1, vbTextCompare) Then
                '工作簿名称是否包含关键词，关键词不区分大小写
                    With GetObject(P & Bookn)
                        For Each Sht In .Worksheets
                            If InStr(1, Sht.Range("a1:b12").Value, Keystr2, vbTextCompare) Then
                            '工作表指定区域内是否包含关键词，关键词不区分大小写
                                If Application.CountIf(Sht.UsedRange, "<>") Then
                                '如果表格存在数据区域
                                    Shtname = Split(Bookn, ".xls")(0) & "-" & Sht.Name
                                    '复制来的工作表以"工作簿-工作表"形式起名。
                                    ThisWorkbook.Sheets(Shtname).Delete
                                    '如果已存在相关表名，则删除
                                    Sht.Copy after:=ThisWorkbook.Worksheets(Sheets.Count)
                                    K = K + 1
                                    '复制Sht到代码所在工作簿所有工作表的后面，并累计个数
                                    ActiveSheet.Name = Shtname
                                    '工作表命名。
                                End If
                            End If
                        Next
                        .Close False '关闭工作簿
                    End With
                End If
            End If
            Bookn = Dir '下一个符合条件的文件
        Loop
        Sh.Select '回到初始工作表
        MsgBox "工作表收集完毕，共收集：" & K & "个"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Sub
    ```

17. 按指定字段将总表数据拆分为多个工作簿

    ```vb
    Sub NewWorkBooks()
        Dim d As Object, arr, brr, r, kr, i&, j&, k&, x&, Mystr$
        Dim Rng As Range, Rg As Range, tRow&, tCol&, aCol&, pd&, mypath$
        Dim Cll As Range, sht As Worksheet
        '第一部分，用户选择保存分表工作簿的路径。
        With Application.FileDialog(msoFileDialogFolderPicker)
       '选择保存工作薄的文件路径
            .AllowMultiSelect = False
            '不允许多选
            If .Show Then
                mypath = .SelectedItems(1)
                '读取选择的文件路径
            Else
                Exit Sub
                '如果没有选择保存路径，则退出程序
            End If
        End With
        If Right(mypath, 1) <> "\" Then mypath = mypath & "\"
        '第二部分遍历总表数据，通过字典将指定字段的不同明细行过滤保存
        Set d = CreateObject("scripting.dictionary") 'set字典
        Set Rg = Application.InputBox("请框选拆分依据列！只能选择单列单元格区域！", Title:="提示", Type:=8)
        '用户选择的拆分依据列
        tCol = Rg.Column '取拆分依据列列标
        tRow = Val(Application.InputBox("请输入总表标题行的行数？"))
        '用户设置总表的标题行数
        If tRow < 0 Then MsgBox "标题行数不能为负数，程序退出。": Exit Sub
        Set Rng = ActiveSheet.UsedRange '总表的数据区域
        Set Cll = ActiveSheet.Cells '用于在分表粘贴和总表同样行高列宽的数据格式
        arr = Rng '数据范围装入数组arr
        tCol = tCol - Rng.Column + 1 '计算依据列在数组中的位置
        aCol = UBound(arr, 2) '数据源的列数
        For i = tRow + 1 To UBound(arr) '遍历数组arr
            If arr(i, tCol) = "" Then arr(i, tCol) = "单元格空白"
            Mystr = arr(i, tCol) '统一转换为字符串格式
            If Not d.exists(Mystr) Then
                d(Mystr) = i '字典中不存在关键词则将行号装入字典
            Else
                d(Mystr) = d(Mystr) & "," & i '如果存在则合并行号，以逗号间隔
            End If
        Next
        '第三部分遍历字典取出分表数据明细，建立不同工作簿保存数据。
        Application.ScreenUpdating = False '关闭屏幕刷新
        Application.DisplayAlerts = False '关闭系统警告信息
        kr = d.keys '字典的key集
        For i = 0 To UBound(kr) '遍历字典key值
            If kr(i) <> "" Then '如果key不为空
                r = Split(d(kr(i)), ",") '取出item里储存的行号
                ReDim brr(1 To UBound(r) + 1, 1 To aCol) '声明放置结果的数组brr
                k = 0
                For x = 0 To UBound(r)
                    k = k + 1 '累加记录行数
                    For j = 1 To aCol '遍历读取列
                        brr(k, j) = arr(r(x), j)
                    Next
                Next
                With Workbooks.Add
                '新建一个工作簿
                    With .Sheets(1).[a1]
                        Cll.Copy '复制粘贴总表的单元格格式
                        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                        Cells.NumberFormat = "@" '设置文本格式，防止文本值变形
                        If tRow > 0 Then .Resize(tRow, aCol) = arr '放标题行
                        .Offset(tRow, 0).Resize(k, aCol) = brr '放置数据区域
                        .Select '激活A1单元格
                    End With
                    .SaveAs mypath & kr(i), xlWorkbookDefault  '保存工作簿
                    .Close True '关闭工作簿
                End With
            End If
        Next
    
        '收
        Set d = Nothing '释放字典
        Erase arr: Erase brr '释放数组
        MsgBox "处理完成。", , "提醒"
        Application.ScreenUpdating = True '恢复屏幕刷新
        Application.DisplayAlerts = True '恢复显示系统警告和消息
    End Sub
    ```
    
18. 常用小代码对工作簿每个工作表进行转置

    ```vb
    Sub transpoose()
    Dim sht As Worksheet, rows As Integer, columns As Integer
    For Each sht In Worksheets
    For rows = 1 To 12
       For columns = 3 To 13
       sht.Cells(columns - 2, rows).Value = sht.Cells(rows, columns).Value
       Next columns
    Next rows
    Next
    End Sub
    ```

19. 常用小代码对工作表非编辑内容进行锁定（code：ccb123456）

    ```vb
    Sub locked_sht()
        Cells.Select
        Selection.Locked = False
        Selection.FormulaHidden = False
        Range("1:14,A15:R1615").Select
        Selection.Locked = True
        Selection.FormulaHidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        ActiveSheet.EnableSelection = xlUnlockedCells
    End Sub
    
    '针对特定报表
    Sub locked_sht()
    Dim IR As Integer
    IR = Sheets("汇总").Range("A65536").End(xlUp).Row
    Sheets("汇总").Unprotect
    Cells.locked = False
    Sheets("汇总").Range("1:12").locked = True
    Sheets("汇总").Range("A11:S" & IR).locked = True
    Sheets("汇总").Protect ("CCB123456")
    Sheets("汇总").EnableSelection = xlUnlockedCells
    End Sub
    
    '针对工作簿内所有工作表
    Sub locked()
    Dim ir As Integer, sht As Worksheet
    For Each sht In Worksheets
        ir = sht.Range("A65536").End(xlUp).Row
        sht.Unprotect
        sht.Cells.locked = False
        sht.Range("1:12").locked = True
        sht.Range("A11:S" & ir).locked = True
        'Range("A10:C78").Select '合并单元格
        'Selection.locked = True '合并单元格
        sht.Protect ("CCB123456")
        sht.EnableSelection = xlUnlockedCells
    Next
    End Sub
    
    '针对文件夹内所有工作簿
    Sub unprotectedbooks()
        Dim mypath$, myname$, sh As Worksheet, ir As Integer
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xls*")
       Do While myname <> ""
        If myname <> ThisWorkbook.Name Then
           With Workbooks.Open(mypath & myname, 0)
               For Each sh In .Sheets
                    ir = sh.Range("A65536").End(xlUp).Row
                    sh.Unprotect
                    sh.Cells.Locked = False
                    sh.Range("1:12").Locked = True
                    sh.Range("A11:S" & ir).Locked = True
                    sh.Protect ("CCB123456")
                    sh.EnableSelection = xlUnlockedCells
               Next
               .Close True
            End With
        End If
        myname = Dir
       Loop
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "finish"
    End Sub
    ```

20. 常用小代码对工作簿中所有的工作表解除保护模式

    ```vb
    Sub unprotected()
    	Dim sht As Worksheet
    	For Each sht In Worksheets
        	sht.Unprotect ("CCB123456") 'CCB123456为密码
    	Next
    End Sub
    ```

21. 常用小代码对工作簿中所有的工作表中的单元格清除格式保留数值

    ```vb
    Sub specialpaste()
    	Dim sht As Worksheet
    	For Each sht In Worksheets
    		sht.Range("a6:i17").Formula = sht.Range("a6:i17").Value
    	Next
    End Sub
    ```

22. 常用小代码对文件夹中所有的工作簿解除保护模式并将公式转为文本

    ```vb
    Sub unprotectedbooks()
        Dim mypath$, myname$, sh As Worksheet
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xlsx")
       Do While myname <> ""
        If myname <> ThisWorkbook.Name Then
           With Workbooks.Open(mypath & myname，0)
               For Each sh In .Sheets
                   sh.Unprotect ("1")
                   sh.Range("a6:i17").Formula = sh.Range("a6:i17").Value
               Next
               .Close True
            End With
        End If
        myname = Dir
       Loop
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "finish"
    End Sub
    ```

23. 常用小代码对工作簿中所有工作表按照颜色排序

    ```vb
    Sub sortbycolors()
        Dim r As Long, sht As Worksheet, i As Integer
        r = Range("a65536").End(xlUp).Row
        Application.ScreenUpdating = False
        For Each sht In Worksheets
            For i = 2 To r
                sht.Cells(i, 2).Value = sht.Cells(i, 1).Interior.ColorIndex 
    'a将颜色的索引值输入到b列
            Next
            sht.Range("a1").CurrentRegion.Sort key1:=sht.Range("b2"), Header:=xlYes 
    '按照b列进行排序
        Next
        Application.ScreenUpdating = True
    End Sub
    ```

24. 常用小代码对工作簿中所有工作表删除空白行

    ```vb
    Sub deleteblankrow()
        Dim firstrow As Long, lastrow As Long, i As Long
        Dim sht As Worksheet
        For Each sht In Worksheets
        firstrow = sht.UsedRange.Row
        lastrow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
                For i = lastrow To firstrow Step -1
                    If Application.WorksheetFunction.CountA(sht.Rows(i)) = 0 Then
                        sht.Rows(i).Delete
                    End If
                Next
         Next
    End Sub
    ```

25. 常用小代码对文件夹中所有工作簿中所有工作表删除空白行

    ```vb
    Sub deleteblankrow()
    Dim mypath$, myname$, sht As Worksheet, firstrow As Long, lastrow As Long, i As Long
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xlsx")
        Do While myname <> ""
            If myname <> ThisWorkbook.Name Then
                With Workbooks.Open(mypath & myname)
                    For Each sht In .Worksheets
                        firstrow = sht.UsedRange.Row
                        lastrow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
                        For i = lastrow To firstrow Step -1
                                If Application.WorksheetFunction.CountA(sht.Rows(i)) = 0 Then
                                    sht.Rows(i).Delete
                                End If
                        Next
                     Next
                     .Close True
                End With
            End If
            myname = Dir
        Loop
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Sub
    ```

26. 常用小代码对工作簿中所有工作表删除空白单元格所在行

    ```vb
    Sub delete_blank_rows()
    Dim sht As Worksheet
    For Each sht In Worksheets
        sht.Range("j:j").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        Next
    End Sub
    
    '对文件夹内工作簿中所有工作表删除空白单元格所在行
    
    Sub delete_blank_rows()
        Dim mypath$, myname$, sh As Worksheet
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xlsx")
       Do While myname <> ""
        If myname <> ThisWorkbook.Name Then
           With Workbooks.Open(mypath & myname, 0)
               For Each sh In .Sheets
                  On Error Resume Next
                  sh.Range("b:b").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
               Next
               .Close True
            End With
        End If
        myname = Dir
       Loop
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "finish"
    End Sub
    ```

27. 常用小代码对文件夹内说有工作簿指定名称的工作表进行汇总

    ```vb
    Sub addsheet()
    Dim mypath$, myname$, sht As Worksheet, firstrow As Long, lastrow As Long, i As Long, k As Integer
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xlsx")
        k = 0
        Do While myname <> ""
            If myname <> ThisWorkbook.Name Then
                With Workbooks.Open(mypath & myname, 0)
                    For Each sht In .Worksheets
                       If sht.Name = "行业表" Then '工作表文件为行业表
                            sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) '将文件加载到第工作表后
                             k = k + 1
                       End If
                     Next
                     .Close True
                End With
            End If
            myname = Dir
        Loop
        MsgBox "工作表收集完毕，共收集：" & k & "个"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Sub
    
    '模糊查询'
    Sub addsheet()
    Dim mypath$, myname$, sht As Worksheet
    Dim firstrow As Long, lastrow As Long, i As Long, k As Integer
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xlsx")
        k = 0
        Do While myname <> ""
            If myname <> ThisWorkbook.Name Then
                With Workbooks.Open(mypath & myname, 0)
                    For Each sht In .Worksheets
                       If InStr(1, sht.Name, "行业", vbTextCompare) Then '工作表文件为包含关键字“行业”
                            sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) '将文件加载到第工作表后
                             k = k + 1
                       End If
                     Next
                     .Close True
                End With
            End If
            myname = Dir
        Loop
        MsgBox "工作表收集完毕，共收集：" & k & "个"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Sub
    
    '对文件夹内所有工作簿添加指定的工作表进行汇总
    Sub addsheet()
    Dim mypath$, myname$, sht As Worksheet, firstrow As Long, lastrow As Long, i As Long, k As Integer
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        mypath = ThisWorkbook.Path & "\"
        myname = Dir(mypath & "*.xls*")
        k = 0
        Do While myname <> ""
            If myname <> ThisWorkbook.Name Then
                With Workbooks.Open(mypath & myname, 0)
                   
                            ThisWorkbook.Worksheets("表1").Copy after:=.Worksheets(ThisWorkbook.Worksheets.Count) '将文件加载到第工作表后
                       
                     .Close True
                End With
            End If
            myname = Dir
        Loop
        MsgBox "完成"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Sub
    ```

28. 常用小代码对文件夹内所有文件复制到指定的子文件中

    ```vb
    Sub FileFunc()
        Dim fso, folder, fc, f1, subfolder, subfolders
        Dim strTmp As String, i As Integer
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder("C:\Users\Administrator\Desktop\VBA\各个分行") '"C:\Users\Administrator\Desktop\VBA\各个分行" the folder's path
        Set fc = folder.Files
        i = 1
        Do While Cells(i, 1) <> "" 'the first column of active sheet should fill the corresponding subfolders name
         For Each f1 In fc
          If InStr(1, f1.Name, Cells(i, 1).Value, vbTextCompare) Then
          fso.CopyFile f1.Path, "C:\Users\Administrator\Desktop\VBA\各个分行\" & Cells(i, 1).Value & "\", True
                                                   'subfolder's path is "C:\Users\Administrator\Desktop\VBA\各个分行\" & Cells(i, 1).Value & "\"
          End If
         Next
        i = i + 1
        Loop
    End Sub
    ```

29. 模糊查询

    ```vb
    Sub timecalculation()
    Dim i As Integer
    	If Not Range("a2").Value Like "yw*" Then
    		MsgBox ("wrong")
    	End If
    End Sub
    ```

30. 将active excel文件复制到对应文件夹下的子文件夹内

    ```vb
    Sub FolderFunc()
        Dim fso, fldr, SubFolders, Subfolder, ftmp1, ftmp2
        Dim Files, File
        Dim strTmp As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists("C:\Users\Administrator\Desktop\新建文件夹") Then
            Set fldr = fso.GetFolder("C:\Users\Administrator\Desktop\新建文件夹")
            Set SubFolders = fldr.SubFolders
            For Each Subfolder In SubFolders
                If Subfolder.Name Like "ha*" Then '可以使用统配符*
                On Error Resume Next
                fso.copyfile "C:\Users\Administrator\Desktop\11.xlsx", "C:\Users\Administrator\Desktop\新建文件夹" & "\" & Subfolder.Name & "\", False '可以使用统配符* fasle表示不覆盖
                End If
            Next
        End If
    End Sub
    ```

31. 将文件下各子文件夹内的excel文件放入指定文件夹内

    ```vb
    Sub FolderFunc()
        Dim fso, fldr, SubFolders, Subfolder, ftmp1, ftmp2
        Dim Files, File
        Dim strTmp As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists("C:\Users\Administrator\Desktop\第一批") Then
            Set fldr = fso.GetFolder("C:\Users\Administrator\Desktop\第一批")
            Set SubFolders = fldr.SubFolders
            For Each Subfolder In SubFolders
            On Error Resume Next
                fso.copyfile "C:\Users\Administrator\Desktop\第一批" & "\" & Subfolder.Name & "\*.xls*", "C:\Users\Administrator\Desktop\新建文件夹", False  '可以使用统配符* fasle表示不覆盖 第二个地址为destination
            Next
        End If
    End Sub
    ```

32. 将对应文件夹下的子文件夹内指定文件进行删除

    ```vb
    Sub FolderFunc()
        Dim fso, fldr, SubFolders, Subfolder, ftmp1, ftmp2
        Dim Files, File
        Dim strTmp As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists("C:\Users\Administrator\Desktop\新建文件夹") Then
            Set fldr = fso.GetFolder("C:\Users\Administrator\Desktop\新建文件夹")
            Set SubFolders = fldr.SubFolders
            For Each Subfolder In SubFolders
            On Error Resume Next
            fso.deletefile "C:\Users\Administrator\Desktop\新建文件夹" & "\" & Subfolder.Name & "\2*.xls" '可以使用统配符* "C:\Users\Administrator\Desktop\新建文件夹" & "\" & Subfolder.Name & "\1*.xls"
            Next
        End If
    End Sub
    ```

33. 将保护文件密码破解

    ```vb
    '方法一'
    Sub RemoveShProtect()
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    Dim i7 As Integer, i8 As Integer, i9 As Integer
    Dim i10 As Integer, i11 As Integer, i12 As Integer
    Dim t As String
    On Error Resume Next
    If ActiveSheet.ProtectContents = False Then
    MsgBox "该工作表没有保护密码!"
    Exit Sub
    End If
    t = Timer
    For i1 = 65 To 66: For i2 = 65 To 66: For i3 = 65 To 66
    For i4 = 65 To 66: For i5 = 65 To 66: For i6 = 65 To 66
    For i7 = 65 To 66: For i8 = 65 To 66: For i9 = 65 To 66
    For i10 = 65 To 66: For i11 = 65 To 66: For i12 = 32 To 126
    ActiveSheet.Unprotect Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) _
    & Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(i10) & Chr(i11) & Chr(i12)
    If ActiveSheet.ProtectContents = False Then
    MsgBox "解除工作表保护!用时" & Format(Timer - t, "0.00") & "秒"
    Exit Sub
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
    End Sub
    '方法二'
    Public Sub 工作表保护密码破解()
    Const DBLSPACE As String = vbNewLine & vbNewLine
    Const AUTHORS As String = DBLSPACE & vbNewLine & _
    "作者:McCormick   JE McGimpsey "
    Const HEADER As String = "工作表保护密码破解"
    Const VERSION As String = DBLSPACE & "版本 Version 1.1.1"
    Const REPBACK As String = DBLSPACE & ""
    Const ZHENGLI As String = DBLSPACE & "                   hfhzi3—戊冥 整理"
    Const ALLCLEAR As String = DBLSPACE & "该工作簿中的工作表密码保护已全部解除!!" & DBLSPACE & "请记得另保存" _
    & DBLSPACE & "注意：不要用在不当地方，要尊重他人的劳动成果！"
    Const MSGNOPWORDS1 As String = "该文件工作表中没有加密"
    Const MSGNOPWORDS2 As String = "该文件工作表中没有加密2"
    Const MSGTAKETIME As String = "解密需花费一定时间,请耐心等候!" & DBLSPACE & "按确定开始破解!"
    Const MSGPWORDFOUND1 As String = "密码重新组合为:" & DBLSPACE & "$$" & DBLSPACE & _
    "如果该文件工作表有不同密码,将搜索下一组密码并修改清除"
    Const MSGPWORDFOUND2 As String = "密码重新组合为:" & DBLSPACE & "$$" & DBLSPACE & _
    "如果该文件工作表有不同密码,将搜索下一组密码并解除"
    Const MSGONLYONE As String = "确保为唯一的?"
    Dim w1 As Worksheet, w2 As Worksheet
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim m As Integer, n As Integer, i1 As Integer, i2 As Integer
    Dim i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
    Dim PWord1 As String
    Dim ShTag As Boolean, WinTag As Boolean
    Application.ScreenUpdating = False
    With ActiveWorkbook
    WinTag = .ProtectStructure Or .ProtectWindows
    End With
    ShTag = False
    For Each w1 In Worksheets
    ShTag = ShTag Or w1.ProtectContents
    Next w1
    If Not ShTag And Not WinTag Then
    MsgBox MSGNOPWORDS1, vbInformation, HEADER
    Exit Sub
    End If
    MsgBox MSGTAKETIME, vbInformation, HEADER
    If Not WinTag Then
    Else
    On Error Resume Next
    Do 'dummy do loop
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    With ActiveWorkbook
    .Unprotect Chr(i) & Chr(j) & Chr(k) & _
    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If .ProtectStructure = False And _
    .ProtectWindows = False Then
    PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
    Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    MsgBox Application.Substitute(MSGPWORDFOUND1, _
    "$$", PWord1), vbInformation, HEADER
    Exit Do 'Bypass all for...nexts
    End If
    End With
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
    Loop Until True
    On Error GoTo 0
    End If
    
    If WinTag And Not ShTag Then
    MsgBox MSGONLYONE, vbInformation, HEADER
    Exit Sub
    End If
    On Error Resume Next
    
    For Each w1 In Worksheets
    'Attempt clearance with PWord1
    w1.Unprotect PWord1
    Next w1
    On Error GoTo 0
    ShTag = False
    For Each w1 In Worksheets
    'Checks for all clear ShTag triggered to 1 if not.
    ShTag = ShTag Or w1.ProtectContents
    Next w1
    If ShTag Then
    For Each w1 In Worksheets
    With w1
    If .ProtectContents Then
    On Error Resume Next
    Do 'Dummy do loop
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    .Unprotect Chr(i) & Chr(j) & Chr(k) & _
    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If Not .ProtectContents Then
    PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
    Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    MsgBox Application.Substitute(MSGPWORDFOUND2, _
    "$$", PWord1), vbInformation, HEADER
    'leverage finding Pword by trying on other sheets
    For Each w2 In Worksheets
    w2.Unprotect PWord1
    Next w2
    Exit Do 'Bypass all for...nexts
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
    Loop Until True
    On Error GoTo 0
    End If
    End With
    Next w1
    End If
    MsgBox ALLCLEAR & AUTHORS & VERSION & REPBACK & ZHENGLI, vbInformation, HEADER
    End Sub
    ```

34. 将保护文件工程密码破解（必须将文件另存为xls模式 同时去除首行 option explicit）

    ```vb
    Private Sub VBAPassword() '你要解保护的Excel文件路径
    Filename = Application.GetOpenFilename("Excel文件（*.xls & *.xla & *.xlt）,*.xls;*.xla;*.xlt", , "VBA破解")
    If Dir(Filename) = "" Then
        MsgBox "没找到相关文件,清重新设置。"
        Exit Sub
        Else
        FileCopy Filename, Filename & ".bak" '备份文件。
    End If
    Dim GetData As String * 5
    Open Filename For Binary As #1
    Dim CMGs As Long
    Dim DPBo As Long
    For i = 1 To LOF(1)
        Get #1, i, GetData
        If GetData = "CMG=""" Then CMGs = i
        If GetData = "[Host" Then DPBo = i - 2: Exit For
    Next
    If CMGs = 0 Then
        MsgBox "请先对VBA编码设置一个保护密码...", 32, "提示"
        Exit Sub
    End If
    Dim St As String * 2
    Dim s20 As String * 1
    '取得一个0D0A十六进制字串
        Get #1, CMGs - 2, St
    '取得一个20十六制字串
        Get #1, DPBo + 16, s20
    '替换加密部份机码
        For i = CMGs To DPBo Step 2
        Put #1, i, St
        Next
    '加入不配对符号
        If (DPBo - CMGs) Mod 2 <> 0 Then
        Put #1, DPBo + 1, s20
        End If
        MsgBox "文件解密成功......", 32, "提示"
    Close #1
    End Sub
    ```

35. 将保护文件密码破解

    ```vb
    Sub FenLei1()'单程序模式
        Dim i As Long, bj As String, rng As Range, sh As Worksheet, r As Long
        Dim i1 As Integer, r1 As Integer, t As Integer, t1 As Integer, s As Integer, i11 As Integer
        Dim c As Integer, c1 As Integer, c2 As Integer
        Dim q As Integer, q1 As Integer
        t1 = ActiveSheet.Range("iv1").End(xlToLeft).Column
        For t = 2 To t1
            For Each sh In Worksheets
            If sh.Name <> "汇总" Then
                For Each rng In sh.Range("A1：K200")
                    If rng.Value Like "*" & Cells(1, t) & "*" Then
                        r = rng.Row
                        i = rng.Column
                        sh.Cells(1, 256 + t) = sh.Cells(r, i + 1)
                    End If
                Next
            End If
            Next
        Next
        
        
        i11 = ActiveSheet.Range("a65535").End(xlUp).Row
        c1 = ActiveSheet.Range("iv1").End(xlToLeft).Column
       
        For i1 = 2 To i11
        For c = 2 To c1
            c2 = Worksheets(Cells(i1, 1).Value).Range("iv1").End(xlToLeft).Column
            Cells(i1, c).Value = Worksheets(Cells(i1, 1).Value).Cells(1, 256 + c).Value
        Next
        Next
        
        For Each sh In Worksheets
        If sh.Name <> "汇总" Then
                sh.Columns("iv:xfd").Clear
        End If
        Next
        
    End Sub
    
    ```
    
36. 冒泡排序

    ```vb
    Sub bubblesort()'按照列排序
        Dim i As Long
        Dim j As Long
        Dim temp As Variant
            For i = 1 To 19 '为计数需要排列指标格式为1开头ubound（array）-1结束
                For j = 5 To 24 - i '为排列指标单元格的列数
                    If Cells(7, j) < Cells(7, j + 1) Then
                        temp = Cells(7, j)
                        Cells(7, j) = Cells(7, j + 1)
                        Cells(7, j + 1) = temp
                    End If
                Next j
            Next i
    End Sub
    
    Sub bubblesort()'按照行排序
        Dim i As Long
        Dim j As Long
        Dim t As Long
        Dim temp As Variant, temp1 As Variant, temp2 As Variant, temp3 As Variant, temp4 As Variant, temp5 As Variant, temp6 As Variant
           For i = 1 To 17
                For j = 1 To 18 - i
                    If Cells(j, 1) < Cells(j + 1, 1) Then
                        temp = Cells(j, 1)
                        Cells(j, 1) = Cells(j + 1, 1)
                        Cells(j + 1, 1) = temp
    
                    End If
                Next j
            Next i
    
    End Sub
    
    '升级版'
    Sub bubblesort()
        Dim i As Long
        Dim j As Long
        Dim t As Long
        Dim temp As Variant, temp1 As Variant, temp2 As Variant, temp3 As Variant, temp4 As Variant, temp5 As Variant, temp6 As Variant
        Dim sh As Worksheet
        For Each sh In Worksheets
            For i = 1 To 19
                For j = 5 To 24 - i
                    If sh.Cells(7, j) < sh.Cells(7, j + 1) Then
                        temp = sh.Cells(7, j)
                        temp1 = sh.Cells(6, j)
                        temp2 = sh.Cells(5, j)
                        temp3 = sh.Cells(4, j)
                        temp4 = sh.Cells(3, j)
                        temp5 = sh.Cells(2, j)
                        temp6 = sh.Cells(1, j)
    
                        sh.Cells(7, j) = sh.Cells(7, j + 1)
                        sh.Cells(6, j) = sh.Cells(6, j + 1)
                        sh.Cells(5, j) = sh.Cells(5, j + 1)
                        sh.Cells(4, j) = sh.Cells(4, j + 1)
                        sh.Cells(3, j) = sh.Cells(3, j + 1)
                        sh.Cells(2, j) = sh.Cells(2, j + 1)
                        sh.Cells(1, j) = sh.Cells(1, j + 1)
    
                        sh.Cells(7, j + 1) = temp
                        sh.Cells(6, j + 1) = temp1
                        sh.Cells(5, j + 1) = temp2
                        sh.Cells(4, j + 1) = temp3
                        sh.Cells(3, j + 1) = temp4
                        sh.Cells(2, j + 1) = temp5
                        sh.Cells(1, j + 1) = temp6
    
                    End If
                Next j
            Next i
        Next
    End Sub
    
    Sub bubblesort() '按照行排序
        Dim i As Long
        Dim j As Long
        Dim t As Long
        Dim z As Long
        Dim rows As Long
        Dim clos As Long
        Dim temp As Variant, temp1 As Variant, temp2 As Variant, temp3 As Variant, temp4 As Variant, temp5 As Variant, temp6 As Variant
        rows = Range("b65536").End(xlUp).Row '表格内部所有行数 需要修改数据开始位置
        clos = Range("b5").End(xlToRight).Column '表格内部所有列数 需要修改数据开始位置
        Debug.Print rows
        ReDim temp(rows - 1 - 4, clos - 1) '数组开始数字为0 如有标题列需要减去标题列rows - 1-标题列
        ReDim temp1(rows - 1 - 4, clos - 1) '数组开始数字为0 如有标题列需要减去标题列rows - 1-标题列
        For i = 0 To rows - 1 - 4 '如有标题列需要减去标题列rows - 1-标题列
            For j = 0 To clos - 1
                temp(i, j) = Cells(i + 1 + 4, j + 1) '如有标题列需要加标题列i + 1 + 4
            Next
        Next
        
        For z = 0 To UBound(temp) - 1
           For t = 0 To UBound(temp) - z - 1
            If temp(t, 8) <= temp(t + 1, 8) Then '需要选择排序列 如 第8列
                For j = 0 To clos - 1
                temp1(t, j) = temp(t, j)
                
                temp(t, j) = temp(t + 1, j)
             
                temp(t + 1, j) = temp1(t, j)
               
                Next
            End If
            Next
        Next
            Range("a5").Resize(rows - 4, clos) = temp '如有标题列需要减去标题列rows-标题列
          
         
    End Sub
    ```

37. 复杂字段中准确找出数值串

    ```vb
    Sub filter_content()
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim t As Integer
    Dim b As Long
        r = [a65536].End(xlUp).Row
        
        For i = 1 To r
            For j = 1 To Len(Cells(i, 1))
              If UCase(Mid(Cells(i, 1), j, 1)) Like "[1234567890]" Then
                 Cells(i, 2) = Cells(i, 2) & Mid(Cells(i, 1).Value, j, 1)
                 Else
                 Cells(i, 2) = Cells(i, 2) & " "
              End If
            Next
         Next
    End Sub
    ```

38. word中寻找特定字符段落

    ```vb
    Sub Macro1()
    Dim wd, mypath$, wj$, i&, x%, zf$
    Dim s As Integer
    Dim nf As Integer
    Set wd = CreateObject("word.application")
    mypath = ThisWorkbook.Path & "\"
    wj = Dir(mypath & "*.doc*")
    Do While wj <> ""
    With wd.Documents.Open(mypath & wj)
        x = .Paragraphs.Count
        Debug.Print x
        For i = x To 1 Step -1
            zf = .Paragraphs(i).Range
            If zf Like "*新发生重大风险事项*" Then
                Debug.Print i
                s = s + 1
                Cells(s + 1, 1) = Split(wj, ".doc*")(0)  '提取文件名
                Cells(s + 1, 2) = .Range(Start:=.Paragraphs(i + 1).Range.Start, End:=.Paragraphs(i + 3).Range.End).Text '提取指定文字后的3段
            End If
    
        Next
        .Close False
    End With
    wj = Dir
    Loop
    wd.Quit
    MsgBox "finish"
    End Sub
    ```

39. word中寻找特定字符段落之间的内容

    ```vb
    Sub Macro()
    Dim wd, mypath$, wj$, i&, x%, p1$, p2$
    Dim s As Integer
    Dim np1 As Integer
    Dim np2 As Integer
    Set wd = CreateObject("word.application")
    mypath = ThisWorkbook.Path & "\"
    wj = Dir(mypath & "*.doc*")
    s = 1
    Do While wj <> ""
    
    With wd.Documents.Open(mypath & wj)
        x = .Paragraphs.Count
        Debug.Print x
        For i = x To 1 Step -1
            p1 = .Paragraphs(i).Range
            If p1 Like "*新暴露纯新不良贷*" Then
                np1 = i
                Debug.Print np1
             End If
        Next
        For i = x To 1 Step -1
            p1 = .Paragraphs(i).Range
            If p1 Like "*新发生重大风险事项*" Then 'x = .Paragraphs(2).Range.Font.Bold判断是否加粗（-1）；x = .Paragraphs(2).Alignment判断段落对齐方式1为居中；x = .Paragraphs(3).Range.Font.Size判断字号
                np2 = i
                Debug.Print np2
             End If
        Next
            Cells(s + 1, 1) = Split(wj, ".doc*")(0)  '提取文件名
            Cells(s + 1, 2) = .Range(Start:=.Paragraphs(np1 + 1).Range.Start, End:=.Paragraphs(np2 - 1).Range.End).Text '提取指定文字后的3段
                
        .Close False
        s = s + 1
    End With
    wj = Dir
    Loop
    wd.Quit
    MsgBox "finish"
    End Sub
    ```

40. word中特定表格寻找特定内容放入汇总表中

    ```vb
    Sub Macro()
    Dim wd, mypath$, wj$, i&, x%, p1$, p2$
    Dim s As Integer
    Dim n As Integer
    Dim np1 As Integer
    Dim np2 As Integer
    Set wd = CreateObject("word.application")
    mypath = ThisWorkbook.Path & "\"
    wj = Dir(mypath & "*.doc*")
    s = 1
    
    Do While wj <> ""
    
    With wd.Documents.Open(mypath & wj)
        x = .tables.Count
        Debug.Print x
        
        For n = 1 To x
        If .tables(n).cell(1, 2).Range.Text Like "*张帅*" Then
        Cells(s + 1, 2) = .tables(n).cell(1, 1)
        Cells(s + 1, 1) = Split(wj, ".doc*")(0)  '提取文件名
        End If
        Next
        .Close False
        s = s + 1
    End With
    wj = Dir
    Loop
    wd.Quit
    MsgBox "finish"
    End Sub
    '方法二 word中特定表格寻找特定内容放入各自表中（需提前建立各个表格）
    Sub Macro()
    Dim wd, mypath$, wj$, i&, x%, p1$, p2$
    Dim s As Integer
    Dim n As Integer
    Dim np1 As Integer
    Dim np2 As Integer
    Dim sht As Worksheet
    Dim str As String
    For Each sht In Worksheets
    str = "*" & sht.Name & "*.doc*"
    Debug.Print str
    Set wd = CreateObject("word.application")
    mypath = ThisWorkbook.Path & "\"
    wj = Dir(mypath & str)
    
    s = 1
    
    
        With wd.Documents.Open(mypath & wj)
            x = .tables.Count
       
        
            For n = 1 To x
                If .tables(n).cell(1, 2).Range.Text Like "*张帅*" Then
                    sht.Cells(s + 1, 2) = .tables(n).cell(1, 1)
                    sht.Cells(s + 1, 1) = Split(wj, ".doc*")(0)  '提取文件名
                End If
            Next
            .Close False
            s = s + 1
        End With
    wj = Dir
    
    wd.Quit
    
    Next
    
    MsgBox "finish"
    End Sub
    '方法三 word中特定表格内容放入各自表中（需提前建立各个表格）
    Sub Macro()
    Dim wd, mypath$, wj$, i&, x%
    Dim s As Integer
    Dim n As Integer
    Dim np1 As Integer
    Dim np2 As Integer
    Dim rw As Integer
    Dim cl As Integer
    Dim sht As Worksheet
    Dim str As String
    For Each sht In Worksheets
    str = "*" & sht.Name & "*.doc*"
    Debug.Print str
    Set wd = CreateObject("word.application")
    mypath = ThisWorkbook.Path & "\"
    wj = Dir(mypath & str)
    s = 1
        With wd.Documents.Open(mypath & wj)
            x = .tables.Count
                For n = 1 To x
                If .tables(n).cell(1, 2).Range.Text Like "*张帅*" Then
                    .tables(n).Range.Copy
                    sht.Range("a1").PasteSpecial xlPasteValues
                End If
            Next
            .Close False
            s = s + 1
        End With
    wj = Dir
    wd.Quit
    Next
    MsgBox "finish"
    End Sub
    ```

41. word中提取特定报表

    ```vb
    Sub WordTabletoExcel()
    Dim WordApp As Object, DOC, mTable, Fn$, Str$
    	On Error Resume Next    '设置容错代码
    	CreateObject("wscript.shell").Run "cmd.exe /c dir """ & ThisWorkbook.Path & "\*.doc"" /s/b>""" & ThisWorkbook.Path & "\list.txt""", False, True     '取得指定目录下的word文档清单
    	Set WordApp = CreateObject("word.application")  '创建word程序项目（用于操作word文档）
    	WordApp.Visible = True  '设定word程序项目可见
    	Open ThisWorkbook.Path & "\list.txt" For Input As #1    '打开清单文件并读取内容
    While Not EOF(1)    '循环读取清单文件各行内容
    	Input #1, Str   '输入一行文本到变量str中
    	If Trim(Str) <> "" Then '如果文本有效则
    		Set DOC = WordApp.Documents.Open(Trim(Str)) '利用word程序项目打开对应的word文档
            With DOC
                For Each mTable In .Tables  '循环文档中的各个表格
                    If Not mTable.Cell(1, 2).Range.Text Like "*张帅*" Then '判断第一行第一列的名称
                        '整个表格复制
                       WordApp.Activate    '激活word程序，使之窗体前置
                        mTable.Range.Copy   '复制表格区域
                        With Windows(1)     '激活excel程序窗体，使之前置
                            .Activate
                            With ThisWorkbook.ActiveSheet   '选中当前使用区A列下面的第一个单元格，并粘贴复制的word中的表格数据
                                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1) = Split(Str, "\")(5)
                                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, 2).Select
                                .Paste
                            End With
                        End With
                '获取某个关键字值
                    End If
                Next mTable
            .Close False    '关闭word文档
            End With
    	End If
    Wend
    Close #1    '关闭清单文件
    If Dir(ThisWorkbook.Path & "\list.txt") <> "" Then Kill ThisWorkbook.Path & "\list.txt"     '删除清单文件
    WordApp.Quit    'word程序项目关闭
    Set DOC = Nothing   '清空对应项目变量
    Set WordApp = Nothing
    End Sub
    ```
    
42. 文件夹中所有word去除空白段落

    ```vb
    Sub WordTabletoExcel()
    
    Dim WordApp As Object, DOC, myParagraph, Fn$, Str$
    On Error Resume Next    '设置容错代码
    CreateObject("wscript.shell").Run "cmd.exe /c dir """ & ThisWorkbook.Path & "\*.doc"" /s/b>""" & ThisWorkbook.Path & "\list.txt""", False, True     '取得指定目录下的word文档清单
    Set WordApp = CreateObject("word.application")  '创建word程序项目（用于操作word文档）
    WordApp.Visible = True  '设定word程序项目可见
    Open ThisWorkbook.Path & "\list.txt" For Input As #1    '打开清单文件并读取内容
    While Not EOF(1)    '循环读取清单文件各行内容
    Input #1, Str   '输入一行文本到变量str中
    If Trim(Str) <> "" Then '如果文本有效则
    Set DOC = WordApp.Documents.Open(Trim(Str)) '利用word程序项目打开对应的word文档
    With DOC
    For Each myParagraph In .Paragraphs
        If Len(Trim(myParagraph.Range)) = 1 Then
           WordApp.Activate    '激活word程序，使之窗体前置
           myParagraph.Range.Delete
        End If
    Next myParagraph
    .Save
    .Close False    '关闭word文档
    End With
    End If
    Wend
    Close #1    '关闭清单文件
    If Dir(ThisWorkbook.Path & "\list.txt") <> "" Then Kill ThisWorkbook.Path & "\list.txt"     '删除清单文件
    WordApp.Quit    'word程序项目关闭
    Set DOC = Nothing   '清空对应项目变量
    Set WordApp = Nothing
    End Sub
    ```

43. 文件夹中所有word指定文本前添加段落内容

    ```vb
    Sub WordTabletoExcel()
    
    Dim WordApp As Object, DOC, myParagraph, Fn$, Str$
    Dim myRange As Range
    On Error Resume Next    '设置容错代码
    CreateObject("wscript.shell").Run "cmd.exe /c dir """ & ThisWorkbook.Path & "\*.doc"" /s/b>""" & ThisWorkbook.Path & "\list.txt""", False, True     '取得指定目录下的word文档清单
    Set WordApp = CreateObject("word.application")  '创建word程序项目（用于操作word文档）
    WordApp.Visible = True  '设定word程序项目可见
    Open ThisWorkbook.Path & "\list.txt" For Input As #1    '打开清单文件并读取内容
    While Not EOF(1)    '循环读取清单文件各行内容
    Input #1, Str   '输入一行文本到变量str中
    If Trim(Str) <> "" Then '如果文本有效则
    Set DOC = WordApp.Documents.Open(Trim(Str)) '利用word程序项目打开对应的word文档
    With DOC
    For Each myParagraph In .Paragraphs
        If myParagraph.Range Like "*1*" Then
                myParagraph.Range.InsertParagraphBefore
                myParagraph.Range.InsertBefore "VBA学习方法" & Chr(10)
        End If
    Next myParagraph
    .Save
    .Close False    '关闭word文档
    End With
    End If
    Wend
    Close #1    '关闭清单文件
    If Dir(ThisWorkbook.Path & "\list.txt") <> "" Then Kill ThisWorkbook.Path & "\list.txt"     '删除清单文件
    WordApp.Quit    'word程序项目关闭
    Set DOC = Nothing   '清空对应项目变量
    Set WordApp = Nothing
    End Sub
    ```

44. 文件夹中所有word指定顺序添加段落内容（需要在excel vbe中引入 microsoft word 15.0 object）

    ```vb
    Sub WordTabletoExcel()
    
    Dim WordApp As Object, DOC, myParagraph, Fn$, Str$
    Dim myRange As Range
    Dim i As Integer
    On Error Resume Next    '设置容错代码
    CreateObject("wscript.shell").Run "cmd.exe /c dir """ & ThisWorkbook.Path & "\1.doc"" /s/b>""" & ThisWorkbook.Path & "\list.txt""", False, True     '取得指定目录下的word文档清单
    Set WordApp = CreateObject("word.application")  '创建word程序项目（用于操作word文档）
    WordApp.Visible = True  '设定word程序项目可见
    Open ThisWorkbook.Path & "\list.txt" For Input As #1    '打开清单文件并读取内容
    While Not EOF(1)    '循环读取清单文件各行内容
    Input #1, Str   '输入一行文本到变量str中
    If Trim(Str) <> "" Then '如果文本有效则
    Set DOC = WordApp.Documents.Open(Trim(Str)) '利用word程序项目打开对应的word文档
    With DOC
    For i = 1 To 5'按照顺序添加段落
    For Each myParagraph In .Paragraphs
        If Len(Trim(myParagraph.Range)) = 1 Then
                myParagraph.Range.InsertBefore Range("a" & i) & ":" & Range("b" & i) & Chr(10)
        End If
    Next myParagraph
    Next
    .Save
    .Close False    '关闭word文档
    End With
    End If
    Wend
    Close #1    '关闭清单文件
    If Dir(ThisWorkbook.Path & "\list.txt") <> "" Then Kill ThisWorkbook.Path & "\list.txt"     '删除清单文件
    WordApp.Quit    'word程序项目关闭
    Set DOC = Nothing   '清空对应项目变量
    Set WordApp = Nothing
    End Sub
    ```

45. 文件夹中所有word指定格式调整内容段落格式（需要在excel vbe中引入 microsoft word 15.0 object）

    ```vb
    Sub WordTabletoExcel_size()
    
    Dim WordApp As Object, DOC, myParagraph, Fn$, Str$
    Dim myRange As Range
    Dim i As Integer
    On Error Resume Next    '设置容错代码
    CreateObject("wscript.shell").Run "cmd.exe /c dir """ & ThisWorkbook.Path & "\1.doc"" /s/b>""" & ThisWorkbook.Path & "\list.txt""", False, True     '取得指定目录下的word文档清单
    Set WordApp = CreateObject("word.application")  '创建word程序项目（用于操作word文档）
    WordApp.Visible = True  '设定word程序项目可见
    Open ThisWorkbook.Path & "\list.txt" For Input As #1    '打开清单文件并读取内容
    While Not EOF(1)    '循环读取清单文件各行内容
    Input #1, Str   '输入一行文本到变量str中
    If Trim(Str) <> "" Then '如果文本有效则
    Set DOC = WordApp.Documents.Open(Trim(Str)) '利用word程序项目打开对应的word文档
    With DOC
    
    For Each myParagraph In .Paragraphs
        If myParagraph.Range Like "zw*" Then
            myParagraph.Range.Font.Name = "彩虹粗仿宋"
            myParagraph.Range.Font.Size = 16
            myParagraph.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            myParagraph.Range.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    
        End If
    Next myParagraph
    
    .Save
    .Close False    '关闭word文档
    End With
    End If
    Wend
    Close #1    '关闭清单文件
    If Dir(ThisWorkbook.Path & "\list.txt") <> "" Then Kill ThisWorkbook.Path & "\list.txt"     '删除清单文件
    WordApp.Quit    'word程序项目关闭
    Set DOC = Nothing   '清空对应项目变量
    Set WordApp = Nothing
    End Sub
    ```

46. word中依据特定段落内容拆分成多个文档（word中vba模块）

    ```vb
    Option Explicit
    
    Sub Word_para_split()
    Dim myrng As Range
    Dim mynm As Range
    Dim arr As Variant
    Dim i As Integer
    Dim j As Integer
    Dim mypath As String
    Dim filename As String
    mypath = "C:\Users\Administrator\Desktop\2\"
    arr = Array(2, 68, 93, 113, 127, 157, 176, 212, 245, 342, 358, 388, 439, 465, 491, 515, 529, 548, 574, 599, 636, 661, 685, 703, 730, 752, 777, 799, 823, 857, 889, 904, 927, 948, 970, 992, 1012, 1033, 1061, 1077, 1103, 1137, 1160, 1194, 1228, 1250, 1284, 1312, 1343, 1366, 1387, 1410, 1491, 1509, 1542, 1570, 1592, 1620, 1658, 1678, 1710, 1737, 1766, 1801, 1818, 1841, 1860, 1896, 1917, 1996, 2091, 2123, 2152, 2174, 2200, 2232, 2263, 2280, 2312, 2336, 2366)
    j = ActiveDocument.Paragraphs.Count
    For i = 0 To UBound(arr) - 1
            Debug.Print i
            Set myrng = ActiveDocument.Range(ActiveDocument.Paragraphs(arr(i) - 1).Range.Start, ActiveDocument.Paragraphs(arr(i + 1) - 2).Range.End)
            Set mynm = ActiveDocument.Paragraphs(arr(i) - 1).Range
            filename = Trim(Replace(mynm.Text, Chr(13), "")) '删除段末的回撤符 chr（10）为换行符
                    
            Debug.Print arr(i)
            Debug.Print filename
            myrng.Select
            Selection.Copy
            Documents.Add
            With ActiveDocument.Content
            .Paste
            End With
            ActiveDocument.SaveAs "C:\Users\Administrator\Desktop\3\" & filename & ".docx"
            ActiveDocument.Close
    Next i
    
    For i = UBound(arr) To UBound(arr)
            Debug.Print i
            Set myrng = ActiveDocument.Range(ActiveDocument.Paragraphs(arr(i) - 1).Range.Start, ActiveDocument.Paragraphs(j).Range.End)
            Set mynm = ActiveDocument.Paragraphs(arr(i) - 1).Range
            filename = Trim(Replace(mynm.Text, Chr(13), ""))
            
            Debug.Print arr(i)
            Debug.Print filename
            myrng.Select
            Selection.Copy
            Documents.Add
            With ActiveDocument.Content
            .Paste
            End With
            ActiveDocument.SaveAs "C:\Users\Administrator\Desktop\3\" & filename & ".docx"
            ActiveDocument.Close
    Next i
    
    End Sub
    ```

47. word删除空白段落（word中vba模块）

    ```vb
    Sub word_del_blank_para()
    Dim myParagraph As Paragraph, n As Integer
    Application.ScreenUpdating = False
    n = 1
    For Each myParagraph In ActiveDocument.Paragraphs
    	If Len(Trim(myParagraph.Range)) = 1 Then
    		myParagraph.Range.Delete
    		n = n + 1
    	End If
    Next
    MsgBox "本次共删除空白段落" & n - 1 & "个"
    Application.ScreenUpdating = True
    End Sub
    ```
    
48. word段落后添加一段新内容（word中vba模块）

    ```vb
    Sub content_Insertafter_para()
    Dim myRange As Range
    Set myRange = ActiveDocument.Paragraphs(1).Range
    	With myRange
    		.InsertAfter "VBA学习方法"
    		.InsertParagraphAfter
    	End With
    End Sub
    ```

    

49. word段落前添加一段新内容（word中vba模块）

    ```vb
    Sub content_InsertBefore_para()
    Dim myRange As Range
    Set myRange = ActiveDocument.Paragraphs(1).Range
    With myRange
    	.InsertParagraphBefore
    	.InsertBefore "VBA学习方法"
    End With
    End Sub
    ```

50. word段落调整样式（word中vba模块）

    ```vb
    Sub word_para_format_setting()
    Dim myParagraph As Paragraph, n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim mypath As String
    Dim MyName As String
    Dim wb
    Application.ScreenUpdating = False
    mypath = ActiveDocument.Path
    MyName = Dir(mypath & "\" & "4*.doc*")
    On Error GoTo ERREXIT
    Do While MyName <> ActiveDocument.Name
    Set wb = Documents.Open(mypath & "\" & MyName)
    i = wb.Paragraphs.Count
    For j = 1 To i
        If wb.Paragraphs(j).Range Like "*三、监测规则*" And j + 2 <= i Then '同时保证j+2不会越界
            wb.Paragraphs(j + 2).Range.Select
             With Selection
                    .Font.Name = "彩虹粗仿宋"
                    .Font.Size = 16
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .ParagraphFormat.OutlineLevel = wdOutlineLevel1
            End With
        End If
    
    Next
    wb.Save
    wb.Close False
    MyName = Dir
    Loop
    Application.ScreenUpdating = True
    ERREXIT:
    Exit Sub '对于找不到的文件直接跳出程序
    End Sub
    ```

    

51. excel中正则表达式挑选数据

    ```vb
    Sub RegExp_Date_Num()
       Dim Res()
       Dim objRegEx As Object
       Dim objMH As Object
       Dim i As Integer
       Dim form As String
       Set objRegEx = CreateObject("vbscript.regexp")
       objRegEx.Pattern = "(\d{4}-\d{2}-\d{2}|\d{4}.\d{2}.\d{2}).*?(([A-Z]{3})*\d+[\d.,]*元)"
       'objRegEx.Pattern = "([\u4E00-\u9FA5]+(省|市|自治区))" '[\u4E00-\u9FA5]+匹配一个或任意个汉字,小括号内的小括号标示子集
       'objRegEx.Pattern = "([\u4E00-\u9FA5]*(省|市|自治区))" '[\u4E00-\u9FA5]*匹配零个或任意个汉字,小括号内的小括号标示子集
       'objRegEx.Pattern = "([\u4E00-\u9FA5、]*(省|市|自治区))" '[\u4E00-\u9FA5]*匹配零个或任意个汉字以及、字符,小括号内的小括号标示子集
       objRegEx.Global = True
       For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
           form = Cells(i, "A")
           Set objMH = objRegEx.Execute(form)
           If objMH.Count > 0 Then
               Cells(i, 2) = CStr(objMH(0).submatches(0)) '查询结果
               Cells(i, 3) = CStr(objMH(0).submatches(1))  '第一个子查询结果
                             'CStr(objMH(j).submatches(2) & objMH(j).submatches(3))'第二个以及第三子查询结果
           End If
       Next
       Set objRegEx = Nothing
       Set objMH = Nothing
    End Sub
    ```

52. excel中正则表达式从本地htm文件中必要数据挑选数据（格式为ANS/ASCII)

    ```vb
    Sub RegExp_Date_Num()
       Dim myfile$, stxt$, mt, n%, arr()
       myfile = ThisWorkbook.Path & "\1.htm"
       Open myfile For Input As #1
       stxt = stxt & StrConv(InputB(LOF(1), 1), vbUnicode)
       Close #1
       With CreateObject("vbscript.regexp")
        .Global = True
        .Pattern = "(\d{4}年\d{2}月\d{2}日\s+\d{2}:\d{2}:\d{2})"
        ReDim arr(1 To .Execute(stxt).Count, 1 To 1)
        Debug.Print .Execute(stxt).Count
        For Each mt In .Execute(stxt)
            n = n + 1
            arr(n, 1) = mt.submatches(0)
        Next
      End With
      Range("a1").Resize(n, 1).Value = arr
    End Sub
    ```

53. excel中正则表达式遍历htm文件中必要数据挑选数据（格式为ANS/ASCII)

    ```vb
    Sub RegExp_Date_Num()
       Dim myfile$, stxt$, mt, n%, arr()
       Dim i As Integer
       Dim t As Integer
       For i = 1 To 3
       myfile = ThisWorkbook.Path & "\新建文件夹\trail\" & i & ".htm"
       Open myfile For Input As #1
       stxt = stxt & StrConv(InputB(LOF(1), 1), vbUnicode)
       Close #1
       With CreateObject("vbscript.regexp")
        .Global = True
        .Pattern = "(\d{4}年\d{2}月\d{2}日\s+\d{2}:\d{2}:\d{2})"
        ReDim arr(1 To .Execute(stxt).Count, 1 To 1)
        Debug.Print .Execute(stxt).Count
        For Each mt In .Execute(stxt)
            n = n + 1
            Cells(n, 2) = mt.submatches(0)
            Cells(n, 1) = i
        Next
      End With
      
      Next i
    End Sub
    ```

54. excel中正则表达式遍历htm文件中必要数据挑选数据（格式为ANS/ASCII)

    ```vb
    Sub RegExp_Date_Num()
       Dim myfile$, stxt$, mt, n%, arr()
       Dim i As Integer
       Dim t As Integer
       For i = 2 To 2
       myfile = ThisWorkbook.Path & "\新建文件夹\trail\" & i & ".htm"
       Open myfile For Input As #1
       stxt = stxt & StrConv(InputB(LOF(1), 1), vbUnicode)
       Close #1
       With CreateObject("vbscript.regexp")
        .Global = True
        .Pattern = "(<TD jQuery\d{13}=.\d{2,}.><SPAN style=.WIDTH: 70px.>[^0-9\s]+[岗|人|管|理|长]</SPAN></TD>)"
        Debug.Print .Execute(stxt).Count
        ReDim arr(1 To .Execute(stxt).Count, 1 To 1)
        Debug.Print .Execute(stxt).Count
        For Each mt In .Execute(stxt)
            n = n + 1
            Cells(n, 3) = mt.submatches(0)
            Cells(n, 4) = i
        Next
      End With
      
      Next i
    End Sub
    ```

55. excel中不打开工作簿的情况下调用工作簿中的宏

    ```vb
    Sub run_prg_withoutopen()
       Workbooks.Open ("C:\Users\Administrator\Desktop\各分行数据\表2-对公分行业.xlsm")
       Application.Run "'表2-对公分行业.xlsm'!aa"
       Workbooks("表2-对公分行业.xlsm").Save
       Workbooks("表2-对公分行业.xlsm").Close
    
    End Sub
    ```

56. excel中限定使用次数的宏

    ```vb
    Private Sub Workbook_Open()'workbook事件中添加
        If Application.UserName <> "张帅" Then
            Call ReadOpenCount
            ThisWorkbook.Save
        End If
    End Sub
    
    Option Explicit'模块中添加
    Sub AddHiddenNames()'编写后需要提前使用
        ThisWorkbook.Names.Add Name:="OpenCount", Visible:=False, RefersTo:="=0"
    End Sub
    
    Sub ReadOpenCount()
    Dim iCount As Byte
    iCount = Evaluate(ThisWorkbook.Names("OpenCount").RefersTo)
    iCount = iCount + 1
    If iCount > 3 Then
        Call KillThisWorkBook
        Else
        ThisWorkbook.Names("OpenCount").RefersTo = "=" & iCount
    End If
    End Sub
    
    Sub KillThisWorkBook()
    With ThisWorkbook
        .Saved = True
        .ChangeFileAccess xlReadOnly
        Kill .FullName
        .Close
    End With
    End Sub
    ```

57. excel中限定使用时间的宏

    ```vb
    Private Sub Workbook_Open()
        If Date >= "2019/12/19" Then
            Application.DisplayAlerts = False
            MsgBox "表格过期"
            With ThisWorkbook
                .Saved = True
                .ChangeFileAccess xlReadOnly
                Kill .FullName
                .Close
            End With
        End If
    End Sub
    ```

58. excel中限定使用次数以及时间的宏

    ```vb
    Private Sub Workbook_Open()
    If Application.UserName <> "张帅" Then
        If Date >= "2019/12/19" Then
            Application.DisplayAlerts = False
            MsgBox "表格过期"
            With ThisWorkbook
                .Saved = True
                .ChangeFileAccess xlReadOnly
                Kill .FullName
                .Close
            End With
        Else
            MsgBox "使用次数监测"
            Call ReadOpenCount'需要在模块中添加相关过程
            ThisWorkbook.Save
        End If
    End If
    End Sub
    
    Option Explicit'模块中添加 不在工作簿事件中添加
    
    Sub AddHiddenNames()'编写后需要提前使用
        ThisWorkbook.Names.Add Name:="OpenCount", Visible:=False, RefersTo:="=0"
    End Sub
    Sub ReadOpenCount()
    Dim iCount As Byte
    iCount = Evaluate(ThisWorkbook.Names("OpenCount").RefersTo)
    iCount = iCount + 1
    If iCount > 3 Then
        Call KillThisWorkBook
        Else
        ThisWorkbook.Names("OpenCount").RefersTo = "=" & iCount
    End If
    End Sub
    Sub KillThisWorkBook()
    With ThisWorkbook
        .Saved = True
        .ChangeFileAccess xlReadOnly
        Kill .FullName
        .Close
    End With
    End Sub
    ```

59. excel中添加聚光灯效果

    ```vb
    Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
        Application.ScreenUpdating = False
            Cells.Interior.ColorIndex = -4142
            '取消单元格原有填充色，但不包含条件格式产生的颜色。
            Rows(Target.Row).Interior.ColorIndex = 33
            '活动单元格整行填充颜色
            Columns(Target.Column).Interior.ColorIndex = 33
            '活动单元格整列填充颜色
        Application.ScreenUpdating = True
    End Sub
    ```

    

