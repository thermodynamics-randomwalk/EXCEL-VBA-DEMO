# EXCEL-ACCESS-SQL

1. ACCESS	基本引用模式

   ```vb
    Sub DoSql_Execute()
       Dim cnn As Object, rst As Object
       Dim Mypath As String, Str_cnn As String, Sql As String
       Dim i As Long
       Set cnn = CreateObject("adodb.connection")
       Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
        Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
       cnn.Open Str_cnn
       Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计
   FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
       Set rst = cnn.Execute(Sql)
       Cells.ClearContents
       For i = 0 To rst.Fields.Count - 1
           Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
       Next
       Range("a2").CopyFromRecordset rst '数据第一行输入位置
       cnn.Close
       Set cnn = Nothing
   End Sub
   
   
    Sub DoSql_Execute()
       Dim cnn As Object, rst As Object
       Dim Mypath As String, Str_cnn As String, Sql As String
       Dim i As Long
       Set cnn = CreateObject("adodb.connection")
       Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
        Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
       cnn.Open Str_cnn
       Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
       Set rst = cnn.Execute(Sql)
        Worksheets("汇总").Cells.ClearContents
       For i = 0 To rst.Fields.Count - 1
           Worksheets("汇总").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
       Next
        Worksheets("汇总").Range("a2").CopyFromRecordset rst '数据第一行输入位置
       cnn.Close
       Set cnn = Nothing
   End Sub
                   
   Sub DoSql_Execute()
       Dim cnn As Object, rst As Object, sht As Worksheet
       Dim Mypath As String, Str_cnn As String, Sql As String
       Dim i As Long
       For Each sht In Worksheets
           If sht.Name = "汇总" Then
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
         Next
   End Sub
   ```

2. if then 模式

   ```vb
   Option Explicit
   Sub DoSql_Execute()
       Dim cnn As Object, rst As Object, sht As Worksheet
       Dim Mypath As String, Str_cnn As String, Sql As String
       Dim i As Long
       For Each sht In Worksheets
           If sht.Name = "汇总" Then
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
           Else
               If sht.Name = "汇总2" Then
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总2").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总2").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总2").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
               Else
                   If sht.Name = "1" Then
                   Set cnn = CreateObject("adodb.connection")
                   Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
                   Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
                   cnn.Open Str_cnn
                   Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
                    Set rst = cnn.Execute(Sql)
                    Worksheets("1").Cells.ClearContents
                    For i = 0 To rst.Fields.Count - 1
                    Worksheets("1").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
                   Next
                   Worksheets("1").Range("a2").CopyFromRecordset rst '数据第一行输入位置
                   cnn.Close
                   Set cnn = Nothing
                   Else
                       If sht.Name = "3" Then
                       Set cnn = CreateObject("adodb.connection")
                       Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
                       Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
                       cnn.Open Str_cnn
                       Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
                       Set rst = cnn.Execute(Sql)
                       Worksheets("3").Cells.ClearContents
                       For i = 0 To rst.Fields.Count - 1
                       Worksheets("3").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
                        Next
                       Worksheets("3").Range("a2").CopyFromRecordset rst '数据第一行输入位置
                       cnn.Close
                       Set cnn = Nothing
                       End If
                   End If
               End If
       End If
       Next
   End Sub
   ```

3. select case 模式

   ```vb
   Sub DoSql_Execute()
       Dim cnn As Object, rst As Object, sht As Worksheet
       Dim Mypath As String, Str_cnn As String, Sql As String
       Dim i As Long
       For Each sht In Worksheets
           Select Case sht.Name
           Case Is = "汇总"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
           Case Is = "汇总2"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总2").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总2").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总2").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
            Case Is = "1"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("1").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("1").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("1").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
          End Select
       Next
   End Sub
   ```

4. select case 模式（一张表输入多个地方sql)

   ```vb
   Sub DoSql_Execute()
       Dim cnn As Object, rst As Object, sht As Worksheet
       Dim Mypath As String, Str_cnn As String, Sql1 As String, Sql As String
       Dim i As Long
       For Each sht In Worksheets
           Select Case sht.Name
           Case Is = "汇总"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总").Range("a:e").ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               
               Sql1 = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql1)
               Worksheets("汇总").Range("i:m").ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 9) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("i2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
            Case Is = "1"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("1").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("1").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("1").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
          End Select
       Next
   End Sub
   
   '方法二
   Sub DoSql_Execute()
   Dim cnn As Object, rst As Object, sht As Worksheet, rst1 As Object
       Dim Mypath As String, Str_cnn As String, Sql1 As String, Sql As String
       Dim i As Long
       For Each sht In Worksheets
           Select Case sht.Name
           Case Is = "汇总"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Sql1 = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst1 = cnn.Execute(Sql1)
               Set rst = cnn.Execute(Sql)
               Worksheets("汇总").Range("i:m").ClearContents
               Worksheets("汇总").Range("a:e").ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               For i = 0 To rst1.Fields.Count - 1
               Worksheets("汇总").Cells(1, i + 9) = rst1.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("汇总").Range("i2").CopyFromRecordset rst1 '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
            Case Is = "1"
               Set cnn = CreateObject("adodb.connection")
               Mypath = ThisWorkbook.Path & "\时间.accdb"  '文件位置
               Str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Mypath  'provider显示为access类型
               cnn.Open Str_cnn
               Sql = "SELECT [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别, Count([201802单笔审批中传统业务类审批].业务编号) AS 业务编号之计数, Sum([201802单笔审批中传统业务类审批].申报金额) AS 申报金额之合计, Sum([201802单笔审批中传统业务类审批].批复金额) AS 批复金额之合计 FROM 201802单笔审批中传统业务类审批 GROUP BY [201802单笔审批中传统业务类审批].客户所属部门, [201802单笔审批中传统业务类审批].审批机构级别"              '//请在此处写入你的SQL代码,表名称可以使用[学生$]
               Set rst = cnn.Execute(Sql)
               Worksheets("1").Cells.ClearContents
               For i = 0 To rst.Fields.Count - 1
               Worksheets("1").Cells(1, i + 1) = rst.Fields(i).Name '设置标题栏位的所在位置
               Next
               Worksheets("1").Range("a2").CopyFromRecordset rst '数据第一行输入位置
               cnn.Close
               Set cnn = Nothing
          End Select
       Next
   End Sub
   
   ```

   