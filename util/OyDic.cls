VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OyDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dic As Object

' ��������
Property Get data() As Variant
    data = dic
End Property

' ��ʼ������
Function init()
    '���ڰ󶨣�����������������������У��Ƽ�
    Set dic = CreateObject("Scripting.Dictionary")
End Function

' �����ַ�
Function setStr(key As String, val As String)
    dic(key) = val
End Function

'��������
Function setInt(key As String, val As Integer)
    dic(key) = val
End Function

' ��ȡֵ
Function getVal(key As String)
    getVal = dic(key)
End Function
