VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����ͬ��"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCalc_Click()

' r�Ǵ������ӵ�ֻ���� c������ֻ��
' feet���ܽ����� heads����ͷ��
' ����r��c����Ϊ����������ʵ�ǲ���Ҫ�ģ���Ϊ���ݺ�����㷨���ǲ����ܲ����������ģ����������͹��ˣ����Ǳ��˵�ֽ��д���Ƿ���������Ҳ�͸���д��������


        Dim r As Single, c As Single
        Dim feet As Integer, heads As Integer

' ��ȡ�����Ľ�����ͷ���� Ҫ��������� ������Ҫ�������������ǼӸ���ʾ��� ������ڶԻ���һ���־�Ҫ���ߣ����Է��ڴ��弤���¼��У�

        feet = Val(TextBox1.Text)
        heads = Val(TextBox2.Text)
        
        
        ' ��������������ʾ������ֻ��
        Debug.Print feet
        Debug.Print heads
        
        


        r = (2 * heads - feet) / 2
        c = heads - r

        If r < 0 Or c < 0 Then
            c = (4 * heads - feet) / 2
            r = heads - c
        End If

        If r < 0 Or c < 0 Then
            MsgBox "�޽Ⱑ��", vbOKOnly, "Ԫ������"
        Else
            TextBox3.Text = c
            TextBox4.Text = r
        End If

End Sub


Private Sub UserForm_Activate()
    MsgBox "�����������ͷ��������д����������Ҫ�ֵ", vbOKOnly, "Ԫ������"

End Sub

