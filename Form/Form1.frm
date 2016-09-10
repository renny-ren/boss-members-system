'本程序共计用时8小时完成  2016年3月7日14:32:01
'第一次使用前需要在注册表的HKEY_CURRENT_USER\Software\VB and VBA Program Settings键值下新建余额1~余额12的数据并初始化为0值，库存1~库存x的数据并初始化，否则会出错



Private Sub Command1_Click()  '保存按钮
Dim j, k As Integer
For j = 0 To 12   '卡号数-1
SaveSetting App.Title, "Set", "余额" & j + 1, Label22(j).Caption  '保存各个余额数据到注册表
SaveSetting App.Title, "Set", "记录" & j + 1, Label14(j).Caption  '保存各个记录数据到注册表
Next
For k = 0 To 16 ' 商品数-1
SaveSetting App.Title, "Set", "库存" & k + 1, Label74(k).Caption  '保存各个库存数据到注册表
Next
End Sub




Private Sub Command2_Click()  '充值按钮
Dim a As String

Select Case Combo1(1).Text
Case Label20(0).Caption: Label22(0).Caption = Val(Label22(0).Caption) + Text2.Text  '加对应余额
a = Label20(0).Caption                                                              'a取值为充值的名字，后面显示提示信息用
Set b = Label22(0)                                                             'b取值为余额，后面显示提示信息用
Case Label20(1).Caption: Label22(1).Caption = Val(Label22(1).Caption) + Text2.Text
a = Label20(1).Caption
Set b = Label22(1)
Case Label20(2).Caption: Label22(2).Caption = Val(Label22(2).Caption) + Text2.Text
a = Label20(2).Caption
Set b = Label22(2)
Case Label20(3).Caption: Label22(3).Caption = Val(Label22(3).Caption) + Text2.Text
a = Label20(3).Caption
Set b = Label22(3)
Case Label20(4).Caption: Label22(4).Caption = Val(Label22(4).Caption) + Text2.Text
a = Label20(4).Caption
Set b = Label22(4)
Case Label20(5).Caption: Label22(5).Caption = Val(Label22(5).Caption) + Text2.Text
a = Label20(5).Caption
Set b = Label22(5)
Case Label20(6).Caption: Label22(6).Caption = Val(Label22(6).Caption) + Text2.Text
a = Label20(6).Caption
Set b = Label22(6)
Case Label20(7).Caption: Label22(7).Caption = Val(Label22(7).Caption) + Text2.Text
a = Label20(7).Caption
Set b = Label22(7)
Case Label20(8).Caption: Label22(8).Caption = Val(Label22(8).Caption) + Text2.Text
a = Label20(8).Caption
Set b = Label22(8)
Case Label20(9).Caption: Label22(9).Caption = Val(Label22(9).Caption) + Text2.Text
a = Label20(9).Caption
Set b = Label22(9)
Case Label20(10).Caption: Label22(10).Caption = Val(Label22(10).Caption) + Text2.Text
a = Label20(10).Caption
Set b = Label22(10)
Case Label20(11).Caption: Label22(11).Caption = Val(Label22(11).Caption) + Text2.Text
a = Label20(11).Caption
Set b = Label22(11)
Case Label20(12).Caption: Label22(12).Caption = Val(Label22(12).Caption) + Text2.Text
a = Label20(12).Caption
Set b = Label22(12)

End Select

If Text2.Text < 20 Then
MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "没有奖励" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
  ElseIf Text2.Text >= 20 And Text2.Text < 50 Then
  b.Caption = b.Caption + 1
  MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励1元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
  
 ElseIf Text2.Text >= 50 And Text2.Text < 100 Then
 b.Caption = b.Caption + 2
 MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励2元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
 
   ElseIf Text2.Text >= 100 And Text2.Text < 200 Then
   b.Caption = b.Caption + 5
   MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励5元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
   
    ElseIf Text2.Text >= 200 Then
    b.Caption = b.Caption + 15
   MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励15元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
   End If


End Sub

Private Sub Command20_Click() '自定义
Select Case Combo1(0).Text

Case Label20(0).Caption: Label22(0).Caption = Label22(0).Caption - Text4.Text
Label14(0).Caption = Label14(0).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - Text4.Text
Label14(1).Caption = Label14(1).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - Text4.Text
Label14(2).Caption = Label14(2).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - Text4.Text
Label14(3).Caption = Label14(3).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - Text4.Text
Label14(4).Caption = Label14(4).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - Text4.Text
Label14(5).Caption = Label14(5).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - Text4.Text
Label14(6).Caption = Label14(6).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - Text4.Text
Label14(7).Caption = Label14(7).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - Text4.Text
Label14(8).Caption = Label14(8).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - Text4.Text
Label14(9).Caption = Label14(9).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - Text4.Text
Label14(10).Caption = Label14(10).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - Text4.Text
Label14(11).Caption = Label14(11).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - Text4.Text
Label14(12).Caption = Label14(12).Caption + Text3.Text & " -" & Text4.Text & "；"

End Select

End Sub





Private Sub Command3_Click()  '瓶装可乐按钮
Select Case Combo1(0).Text

Case Label20(0).Caption: Label22(0).Caption = Label22(0).Caption - 2.8   '2.8为该商品价格
Label14(0).Caption = Label14(0).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.8
Label14(1).Caption = Label14(1).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.8
Label14(2).Caption = Label14(2).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.8
Label14(3).Caption = Label14(3).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.8
Label14(4).Caption = Label14(4).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.8
Label14(5).Caption = Label14(5).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.8
Label14(6).Caption = Label14(6).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.8
Label14(7).Caption = Label14(7).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.8
Label14(8).Caption = Label14(8).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.8
Label14(9).Caption = Label14(9).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.8
Label14(10).Caption = Label14(10).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.8
Label14(11).Caption = Label14(11).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.8
Label14(12).Caption = Label14(12).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

End Select

End Sub

Private Sub Command4_Click()  '方便面按钮
Select Case Combo1(0).Text

Case Label20(0).Caption: Label22(0).Caption = Label22(0).Caption - 3.8
Label14(0).Caption = Label14(0).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.8
Label14(1).Caption = Label14(1).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.8
Label14(2).Caption = Label14(2).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.8
Label14(3).Caption = Label14(3).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.8
Label14(4).Caption = Label14(4).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.8
Label14(5).Caption = Label14(5).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.8
Label14(6).Caption = Label14(6).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.8
Label14(7).Caption = Label14(7).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.8
Label14(8).Caption = Label14(8).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.8
Label14(9).Caption = Label14(9).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.8
Label14(10).Caption = Label14(10).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.8
Label14(11).Caption = Label14(11).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.8
Label14(12).Caption = Label14(12).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
End Select

End Sub

Private Sub Command5_Click() '火腿肠按钮
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 0.8  '扣除金额
Label14(0).Caption = Label14(0).Caption + "火腿肠 -0.8；"                '添加记录
Label74(2).Caption = Label74(2).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 0.8
Label14(1).Caption = Label14(1).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 0.8
Label14(2).Caption = Label14(2).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 0.8
Label14(3).Caption = Label14(3).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 0.8
Label14(4).Caption = Label14(4).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 0.8
Label14(5).Caption = Label14(5).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 0.8
Label14(6).Caption = Label14(6).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 0.8
Label14(7).Caption = Label14(7).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 0.8
Label14(8).Caption = Label14(8).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 0.8
Label14(9).Caption = Label14(9).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 0.8
Label14(10).Caption = Label14(10).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 0.8
Label14(11).Caption = Label14(11).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 0.8
Label14(12).Caption = Label14(12).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
End Select
End Sub
Private Sub Command8_Click()  '劲仔
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1  '扣除金额
Label14(0).Caption = Label14(0).Caption + "劲仔-1；"                '添加记录
Label74(3).Caption = Label74(3).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1
Label14(1).Caption = Label14(1).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1
Label14(2).Caption = Label14(2).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1
Label14(3).Caption = Label14(3).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1
Label14(4).Caption = Label14(4).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1
Label14(5).Caption = Label14(5).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1
Label14(6).Caption = Label14(6).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1
Label14(7).Caption = Label14(7).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1
Label14(8).Caption = Label14(8).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1
Label14(9).Caption = Label14(9).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1
Label14(10).Caption = Label14(10).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1
Label14(11).Caption = Label14(11).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1
Label14(12).Caption = Label14(12).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
End Select

End Sub
Private Sub Command9_Click()  '土老帽火腿香干
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1.9  '扣除金额
Label14(0).Caption = Label14(0).Caption + "土老帽-1.9；"                '添加记录
Label74(4).Caption = Label74(4).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1.9
Label14(1).Caption = Label14(1).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1.9
Label14(2).Caption = Label14(2).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1.9
Label14(3).Caption = Label14(3).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1.9
Label14(4).Caption = Label14(4).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1.9
Label14(5).Caption = Label14(5).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1.9
Label14(6).Caption = Label14(6).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1.9
Label14(7).Caption = Label14(7).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1.9
Label14(8).Caption = Label14(8).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1.9
Label14(9).Caption = Label14(9).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1.9
Label14(10).Caption = Label14(10).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1.9
Label14(11).Caption = Label14(11).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1.9
Label14(12).Caption = Label14(12).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
End Select

End Sub
Private Sub Command10_Click()  '烧烤素鸡
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.3  '扣除金额
Label14(0).Caption = Label14(0).Caption + "素鸡-2.3；"                '添加记录
Label74(5).Caption = Label74(5).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.3
Label14(1).Caption = Label14(1).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.3
Label14(2).Caption = Label14(2).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.3
Label14(3).Caption = Label14(3).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.3
Label14(4).Caption = Label14(4).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.3
Label14(5).Caption = Label14(5).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.3
Label14(6).Caption = Label14(6).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.3
Label14(7).Caption = Label14(7).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.3
Label14(8).Caption = Label14(8).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.3
Label14(9).Caption = Label14(9).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.3
Label14(10).Caption = Label14(10).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.3
Label14(11).Caption = Label14(11).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.3
Label14(12).Caption = Label14(12).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
End Select

End Sub
Private Sub Command11_Click()  '乐事薯片
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.8  '扣除金额
Label14(0).Caption = Label14(0).Caption + "薯片-3.8；"                '添加记录
Label74(6).Caption = Label74(6).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.8
Label14(1).Caption = Label14(1).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.8
Label14(2).Caption = Label14(2).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.8
Label14(3).Caption = Label14(3).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.8
Label14(4).Caption = Label14(4).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.8
Label14(5).Caption = Label14(5).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.8
Label14(6).Caption = Label14(6).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.8
Label14(7).Caption = Label14(7).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.8
Label14(8).Caption = Label14(8).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.8
Label14(9).Caption = Label14(9).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.8
Label14(10).Caption = Label14(10).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.8
Label14(11).Caption = Label14(11).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.8
Label14(12).Caption = Label14(12).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
End Select

End Sub
Private Sub Command12_Click()  '乐事薯片烤肉味
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.5  '扣除金额
Label14(0).Caption = Label14(0).Caption + "薯片扁的-3.5；"                '添加记录
Label74(7).Caption = Label74(7).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.5
Label14(1).Caption = Label14(1).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.5
Label14(2).Caption = Label14(2).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.5
Label14(3).Caption = Label14(3).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.5
Label14(4).Caption = Label14(4).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.5
Label14(5).Caption = Label14(5).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.5
Label14(6).Caption = Label14(6).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.5
Label14(7).Caption = Label14(7).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.5
Label14(8).Caption = Label14(8).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.5
Label14(9).Caption = Label14(9).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.5
Label14(10).Caption = Label14(10).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.5
Label14(11).Caption = Label14(11).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.5
Label14(12).Caption = Label14(12).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
End Select

End Sub
Private Sub Command13_Click()  '卫龙大面筋
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.5  '扣除金额
Label14(0).Caption = Label14(0).Caption + "大面筋-2.5；"                '添加记录
Label74(8).Caption = Label74(8).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.5
Label14(1).Caption = Label14(1).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.5
Label14(2).Caption = Label14(2).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.5
Label14(3).Caption = Label14(3).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.5
Label14(4).Caption = Label14(4).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.5
Label14(5).Caption = Label14(5).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.5
Label14(6).Caption = Label14(6).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.5
Label14(7).Caption = Label14(7).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.5
Label14(8).Caption = Label14(8).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.5
Label14(9).Caption = Label14(9).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.5
Label14(10).Caption = Label14(10).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.5
Label14(11).Caption = Label14(11).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.5
Label14(12).Caption = Label14(12).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
End Select

End Sub
Private Sub Command14_Click()  '小米锅巴
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.3  '扣除金额
Label14(0).Caption = Label14(0).Caption + "小米锅巴-2.3；"                '添加记录
Label74(9).Caption = Label74(9).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.3
Label14(1).Caption = Label14(1).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.3
Label14(2).Caption = Label14(2).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.3
Label14(3).Caption = Label14(3).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.3
Label14(4).Caption = Label14(4).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.3
Label14(5).Caption = Label14(5).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.3
Label14(6).Caption = Label14(6).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.3
Label14(7).Caption = Label14(7).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.3
Label14(8).Caption = Label14(8).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.3
Label14(9).Caption = Label14(9).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.3
Label14(10).Caption = Label14(10).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.3
Label14(11).Caption = Label14(11).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.3
Label14(12).Caption = Label14(12).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
End Select

End Sub
Private Sub Command15_Click()  '奇多干杯脆
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1.7  '扣除金额
Label14(0).Caption = Label14(0).Caption + "干杯脆-1.7；"                '添加记录
Label74(10).Caption = Label74(10).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1.7
Label14(1).Caption = Label14(1).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1.7
Label14(2).Caption = Label14(2).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1.7
Label14(3).Caption = Label14(3).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1.7
Label14(4).Caption = Label14(4).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1.7
Label14(5).Caption = Label14(5).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1.7
Label14(6).Caption = Label14(6).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1.7
Label14(7).Caption = Label14(7).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1.7
Label14(8).Caption = Label14(8).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1.7
Label14(9).Caption = Label14(9).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1.7
Label14(10).Caption = Label14(10).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1.7
Label14(11).Caption = Label14(11).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1.7
Label14(12).Caption = Label14(12).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
End Select

End Sub
Private Sub Command16_Click()  '好丽友
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1.4  '扣除金额
Label14(0).Caption = Label14(0).Caption + "好丽友-1.4；"                '添加记录
Label74(11).Caption = Label74(11).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1.4
Label14(1).Caption = Label14(1).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1.4
Label14(2).Caption = Label14(2).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1.4
Label14(3).Caption = Label14(3).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1.4
Label14(4).Caption = Label14(4).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1.4
Label14(5).Caption = Label14(5).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1.4
Label14(6).Caption = Label14(6).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1.4
Label14(7).Caption = Label14(7).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1.4
Label14(8).Caption = Label14(8).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1.4
Label14(9).Caption = Label14(9).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1.4
Label14(10).Caption = Label14(10).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1.4
Label14(11).Caption = Label14(11).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1.4
Label14(12).Caption = Label14(12).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
End Select

End Sub

Private Sub Command17_Click() '旺旺小小酥
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.3  '扣除金额
Label14(0).Caption = Label14(0).Caption + "小小酥-3.3；"                '添加记录
Label74(12).Caption = Label74(12).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.3
Label14(1).Caption = Label14(1).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.3
Label14(2).Caption = Label14(2).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.3
Label14(3).Caption = Label14(3).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.3
Label14(4).Caption = Label14(4).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.3
Label14(5).Caption = Label14(5).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.3
Label14(6).Caption = Label14(6).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.3
Label14(7).Caption = Label14(7).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.3
Label14(8).Caption = Label14(8).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.3
Label14(9).Caption = Label14(9).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.3
Label14(10).Caption = Label14(10).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.3
Label14(11).Caption = Label14(11).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.3
Label14(12).Caption = Label14(12).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
End Select

End Sub
Private Sub Command18_Click()  '浪味仙
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.8  '扣除金额
Label14(0).Caption = Label14(0).Caption + "浪味仙-2.8；"                '添加记录
Label74(13).Caption = Label74(13).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.8
Label14(1).Caption = Label14(1).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.8
Label14(2).Caption = Label14(2).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.8
Label14(3).Caption = Label14(3).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.8
Label14(4).Caption = Label14(4).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.8
Label14(5).Caption = Label14(5).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.8
Label14(6).Caption = Label14(6).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.8
Label14(7).Caption = Label14(7).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.8
Label14(8).Caption = Label14(8).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.8
Label14(9).Caption = Label14(9).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.8
Label14(10).Caption = Label14(10).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.8
Label14(11).Caption = Label14(11).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.8
Label14(12).Caption = Label14(12).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
End Select

End Sub

Private Sub Command19_Click() '听装可乐
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.1  '扣除金额
Label14(0).Caption = Label14(0).Caption + "听可乐-2.1；"                '添加记录
Label74(14).Caption = Label74(14).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.1
Label14(1).Caption = Label14(1).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.1
Label14(2).Caption = Label14(2).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.1
Label14(3).Caption = Label14(3).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.1
Label14(4).Caption = Label14(4).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.1
Label14(5).Caption = Label14(5).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.1
Label14(6).Caption = Label14(6).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.1
Label14(7).Caption = Label14(7).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.1
Label14(8).Caption = Label14(8).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.1
Label14(9).Caption = Label14(9).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.1
Label14(10).Caption = Label14(10).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.1
Label14(11).Caption = Label14(11).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.1
Label14(12).Caption = Label14(12).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
End Select
End Sub
Private Sub Command21_Click() '八宝粥
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.4  '扣除金额
Label14(0).Caption = Label14(0).Caption + "八宝粥-3.4；"                '添加记录
Label74(15).Caption = Label74(15).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.4
Label14(1).Caption = Label14(1).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.4
Label14(2).Caption = Label14(2).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.4
Label14(3).Caption = Label14(3).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.4
Label14(4).Caption = Label14(4).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.4
Label14(5).Caption = Label14(5).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.4
Label14(6).Caption = Label14(6).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.4
Label14(7).Caption = Label14(7).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.4
Label14(8).Caption = Label14(8).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.4
Label14(9).Caption = Label14(9).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.4
Label14(10).Caption = Label14(10).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.4
Label14(11).Caption = Label14(11).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.4
Label14(12).Caption = Label14(12).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
End Select

End Sub

Private Sub Command22_Click()  '百醇
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 6.2  '扣除金额
Label14(0).Caption = Label14(0).Caption + "百醇-6.2；"                '添加记录
Label74(16).Caption = Label74(16).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 6.2
Label14(1).Caption = Label14(1).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 6.2
Label14(2).Caption = Label14(2).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 6.2
Label14(3).Caption = Label14(3).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 6.2
Label14(4).Caption = Label14(4).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 6.2
Label14(5).Caption = Label14(5).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 6.2
Label14(6).Caption = Label14(6).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 6.2
Label14(7).Caption = Label14(7).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 6.2
Label14(8).Caption = Label14(8).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 6.2
Label14(9).Caption = Label14(9).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 6.2
Label14(10).Caption = Label14(10).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 6.2
Label14(11).Caption = Label14(11).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 6.2
Label14(12).Caption = Label14(12).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
End Select
End Sub
Private Sub Command6_Click() ' 添加姓名按钮
a = GetSetting(App.Title, "Set", "变量a")  '从注册表取得变量a的值，用于判断下一个添加的姓名是第几个
Load Label20(a + 1)   '加载下一个姓名
Label20(a + 1).Top = Label20(a).Top + 500 '调整位置
Label20(a + 1).Visible = True
Label20(a + 1).Caption = Text1.Text   '设置姓名
SaveSetting App.Title, "Set", "姓名" & (a + 2), Label20(a + 1).Caption  '保存姓名的数据到注册表，下一次加载窗体时读取
a = a + 1
SaveSetting App.Title, "Set", "变量a", a
Combo1(0).AddItem Text1.Text
Combo1(1).AddItem Text1.Text

End Sub

Private Sub Command7_Click()
For i = 0 To 12          '卡号数-1
Label14(i).Caption = ""
Next
End Sub


Private Sub Form_Load()
Dim j, k As Integer
For j = 0 To 12         '卡号数-1
Label22(j).Caption = GetSetting(App.Title, "Set", "余额" & j + 1)   '读取余额数据
Label14(j).Caption = GetSetting(App.Title, "Set", "记录" & j + 1)   '读取记录数据
Next
    
For k = 0 To 16         '商品数-1
Label74(k).Caption = GetSetting(App.Title, "Set", "库存" & k + 1)   '读取库存数据
Next

  a = GetSetting(App.Title, "Set", "变量a")
If a <> 0 Then
Dim i As Integer
i = 0
  While i <> a
  Load Label20(i + 1)
  Label20(i + 1).Top = Label20(i).Top + 500
  Label20(i + 1).Visible = True
  Label20(i + 1).Caption = GetSetting(App.Title, "Set", "姓名" & (i + 2))   '读取姓名数据
  Combo1(0).AddItem Label20(i + 1).Caption
  Combo1(1).AddItem Label20(i + 1).Caption
  i = i + 1
  Wend
End If
 
End Sub

