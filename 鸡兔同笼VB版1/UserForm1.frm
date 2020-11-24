VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "鸡兔同笼"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCalc_Click()

' r是代表兔子的只数， c代表鸡的只数
' feet是总脚数， heads是总头数
' 这里r和c声明为符点数，其实是不必要的，因为根据后面的算法，是不可能产生符点数的，所以整数就够了，但是别人的纸上写的是符点数，我也就跟着写符点数吧


        Dim r As Single, c As Single
        Dim feet As Integer, heads As Integer

' 读取给定的脚数和头数。 要给正整数额， ！！不要给负整数，还是加个提示框吧 （这个在对话框一出现就要告诉，所以放在窗体激活事件中）

        feet = Val(TextBox1.Text)
        heads = Val(TextBox2.Text)
        
        
        ' 在立即窗口中显示脚数和只数
        Debug.Print feet
        Debug.Print heads
        
        


        r = (2 * heads - feet) / 2
        c = heads - r

        If r < 0 Or c < 0 Then
            c = (4 * heads - feet) / 2
            r = heads - c
        End If

        If r < 0 Or c < 0 Then
            MsgBox "无解啊哈", vbOKOnly, "元月提醒"
        Else
            TextBox3.Text = c
            TextBox4.Text = r
        End If

End Sub


Private Sub UserForm_Activate()
    MsgBox "请输入脚数和头数，请填写正整数，不要填负值", vbOKOnly, "元月提醒"

End Sub

