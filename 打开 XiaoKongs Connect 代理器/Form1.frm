VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Connect"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4125
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4125
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Me.Hide
Call InternetSetOption(0, 39, 0, 0) 'д��ע������ã����ɲ�����IEʵ�������ַ�ĸ���
MsgBox " KO Connect", vbInformation, "XIAOKONGS"
End
End Sub
