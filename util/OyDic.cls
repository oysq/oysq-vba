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

' 返回数据
Property Get data() As Variant
    data = dic
End Property

' 初始化数组
Function init()
    '后期绑定：方便代码在其他电脑上运行，推荐
    Set dic = CreateObject("Scripting.Dictionary")
End Function

' 设置字符
Function setStr(key As String, val As String)
    dic(key) = val
End Function

'设置数字
Function setInt(key As String, val As Integer)
    dic(key) = val
End Function

' 获取值
Function getVal(key As String)
    getVal = dic(key)
End Function

