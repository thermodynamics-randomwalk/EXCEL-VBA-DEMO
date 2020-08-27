# EXCEL-LOTUS

## 背对背发送邮件

```vb
Sub back_to_back_email_send()
Dim no As Object
Dim db As Object
Dim doc As Object
Dim field As Object
Dim z As Integer
Set no = CreateObject("Notes.NotesSession")
Set db = no.CURRENTDATABASE

For z = 1 To 38
Set doc = db.createdocument
Set field = doc.CREATERICHTEXTITEM("Body")

With field
.appendtext "各省、自治区、直辖市分行，总行直属分行，苏州分行：" & vbCrLf & "     依据建信审〔2019〕57号文中关于填报授信审批工作季度报表的相关要求，总行通过CMISII系统，现将疑似还款来源为地方财政资金的授信业务审批明细（附表5）发送各分行。请各分行依照附表所列项目清单，将确属于依赖地方财政资金还款的授信业务具体情况填报完整；同时，将分行掌握的、未在附表中列示的还款来源依赖地方财政资金的授信业务一并填入报表，确保报表数据的完整性。" & vbCrLf & _
"     针对境内外各机构汇总的《全球授信完成情况统计表》（季报）中存在的问题，再次重申相关要求：'完成全球授信'指在统计期内形成最终批复结论的全球授信申报，不含申报流程中尚未批复的项目;该表仅统计填报分行作为牵头行的情况，作为成员行的不需要填报;由于该表为季度统计汇总，因此表中数据应为该季度内的完成情况，不属于本季度的应及时删除，同时，文字总结中对全球授信完成情况统计也应为季度维度数据，不应包含以往历史数据；全球授信统计口径为涵盖海外需求的全部综合授信业务，包括：海外机构已形成实际占用的，为海外机构预留额度的，以及海外机构有营销需求的各类潜在客户等，各行应进一步明确该统计口径，主动了解海外机构需求，准确在表格中填写符合以上条件的海外参与机构名称，确保统计口径的一致性，提高统计数据的有效性。" & vbCrLf & _
"     本次季度报告上报截止日期为2019年1月3日，请各分行认真按照建信审〔2019〕57号文中的相关要求，做好相关工作，及时上报总行，确保上报工作报告及报表的质量和效率。" & vbCrLf & _
"     联系人： 张帅 电话：010-67596629   邮箱：zhangshuai.zh@ccb.com"
.addnewline 2
.embedobject 1454, "", Sheets(1).Cells(z, 1)
End With

With doc
.form = "memo"
.sendto = CStr(Sheets(1).Cells(z, 2))
.Subject = "关于按附表所列项目清单填报建信审〔2019〕57号文中表5以及重申《全球授信完成情况统计表》报表填报要求的通知"
.savemessageonsend = True
.posteddate = Now()
.send 0
End With

Set doc = Nothing
Set field = Nothing

Next z

End Sub
```

## 背对背发送邮件 带回执模式

```vb
Sub back_to_back_email_send()
Dim no As Object
Dim db As Object
Dim doc As Object
Dim field As Object
Dim z As Integer
Set no = CreateObject("Notes.NotesSession")
Set db = no.CURRENTDATABASE
For z = 1 To 10
Set doc = db.createdocument
Set field = doc.CREATERICHTEXTITEM("Body")

With field

.appendtext "各海外机构：" & vbCrLf & "     为减轻海外手工报送报表工作负担，尽快实现新一代核心系统数据自动提取，近年来总行持续推动“新一代”项目组开展海外机构授信业务审批明细报表开发上线工作。近期，该报表已开发完毕并上线运行。现将从系统抽取的“各海外机构2018年全年授信业务审批明细报表”下发各海外机构，请认真对照自身手工审批台账，逐笔核对报表数据的完整性、准确性，梳理存在的问题并报告总行，以便总行会同技术部门持续改进报表质量。" & vbCrLf & "     请重点核对以下问题，并在附表中标注说明（涂黄部分为必须填写内容）:" & vbCrLf & _
"      1、对于丢失的审批记录，核实丢失原因属于相关业务在“线下”审批（包括在其他系统、未在“新一代”系统审批的情况），还是属于在统计过程中系统取数丢失的情况；" & vbCrLf & _
"      2、对于在源系统有审批记录，但统计取数丢失的业务，进一步核实丢失业务笔数总量，是否具有共同特征（比如均属于否决项目，均属于同一授信产品等），同时，根据手工台账，补齐明细数据并单独标注；" & vbCrLf & _
"      3、对于字段值为空的审批记录，核实为空的原因属于相关数值在源生产系统没有录入，还是在源生产系统已经录入但统计过程中取数丢失；" & vbCrLf & _
"      4、核对发现的其他报表数据错误。" & vbCrLf & _
"      根据总行统一工作安排，系统自动抽取的报表数据将作为总行实施条线业务管理、开展年终考核评价的主要决策依据，因此，请各海外机构高度重视此次数据核对工作，仔细查找问题原因，并及时向总行报告情况。核对标注后的明细表请于1月25日（周五下班前）反馈我处。" & vbCrLf & _
"      联系人： 宋旭群  67596548  邮箱：songxuqun/zh/ccb  张帅 电话：010-67596629   邮箱：zhangshuai.zh@ccb.com"
.addnewline 2
.embedobject 1454, "", Sheets(1).Cells(z, 1)
End With

With doc
.form = "memo"
.ReturnReceipt = "1"

.sendto = CStr(Sheets(1).Cells(z, 2))
.Subject = "关于认真核对系统统计报表质量的通知"
.savemessageonsend = True
.posteddate = Now()
.send 0
End With

Set doc = Nothing
Set field = Nothing

Next z

End Sub

```

## 背对背发送邮件 发送多个附件

```vb
Sub send_with_lotus_attachments()
Dim noSession As Object, noDatabase As Object
Dim noDocument As Object, noAttachment As Object
Dim vaFiles As Variant
Dim i As Long
Const EMBED_ATTACHMENT = 1454
Const stSubject As String = "for lotus VBA programming test only"
Const stMsg As String = "this file is for you!"
Dim vaRecipient As Variant
vaRecipient = VBA.Array("zhangshuai/zh/ccb@ccb", "zhangshuai1990616@163.com")
vaFiles = Application.GetOpenFilename(filefilter:="", Title:="attach files for outgoing", MultiSelect:=True)
If Not IsArray(vaFiles) Then Exit Sub
Set noSession = CreateObject("Notes.NotesSession")
Set noDatabase = noSession.CURRENTDATABASE
If noDatabase.IsOpen = False Then noDatabase.openmail
Set noDocument = noDatabase.CREATEDOCUMENT
Set noAttachment = noDocument.CREATERICHTEXTITEM("Body")
With noAttachment
For i = 1 To UBound(vaFiles)
.embedobject EMBED_ATTACHMENT, "", vaFiles(i)
Next i
End With
With noDocument
.form = "memo"
.sendto = vaRecipient
.Subject = stSubject
.body = stMsg
.savemessageonsend = True
.posteddate = Now()
.send 0, vaRecipient
End With
    Set noDocument = Nothing
    Set noDatabase = Nothing
    Set noSession = Nothing
MsgBox "this file is send ok", vbInformation
End Sub
```

## 背对背发送邮件 发送多个附件 getdatabase版本

```vb
Sub sendwithlotus()
Dim noSession As Object, noDatabase As Object
Dim noDocument As Object, noAttachment As Object
Dim vaFiles As Variant
Dim i As Long
Const EMBED_ATTACHMENT = 1454
Const stSubject As String = "for lotus VBA programming test only"
Const stMsg As String = "this file is for you!"
Dim vaRecipient As Variant
vaRecipient = VBA.Array("zhangshuai/zh/ccb@ccb")
vaFiles = Application.GetOpenFilename(filefilter:="", Title:="attach files for outgoing", MultiSelect:=True)
If Not IsArray(vaFiles) Then Exit Sub
Set noSession = CreateObject("Notes.NotesSession")
Set noDatabase = noSession.getdatabase("", "D:\zhangshuai.nsf")
If noDatabase.IsOpen = False Then noDatabase.openmail
Set noDocument = noDatabase.CREATEDOCUMENT
Set noAttachment = noDocument.CREATERICHTEXTITEM("Body")
With noAttachment
For i = 1 To UBound(vaFiles)
.embedobject EMBED_ATTACHMENT, "", vaFiles(i)
Next i
End With
With noDocument
.form = "memo"
.sendto = vaRecipient
.Subject = stSubject
.body = stMsg
.savemessageonsend = True
.posteddate = Now()
.send 0, vaRecipient
End With
    Set noDocument = Nothing
    Set noDatabase = Nothing
    Set noSession = Nothing
MsgBox "this file is send ok", vbInformation
End Sub
```

