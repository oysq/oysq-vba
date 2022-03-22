VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OyArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private arr() As Variant

' 返回数据
Property Get data() As Variant
    data = arr
End Property

' 初始化数组
Function init()
    ReDim arr(0)
End Function

'添加字符
Function addStr(str As String)
    Dim i As Integer
    i = UBound(arr)
    If i = 0 And IsEmpty(arr(i)) Then
        arr(i) = str
    Else
        ReDim Preserve arr(i + 1)
        arr(i + 1) = str
    End If
End Function

'添加数字
Function addInt(num As Integer)
    Dim i As Integer
    i = UBound(arr)
    If i = 0 And IsEmpty(arr(i)) Then
        arr(i) = num
    Else
        ReDim Preserve arr(i + 1)
        arr(i + 1) = num
    End If
End Function

' 打印一维数组
Function printArr()
    Debug.Print "--- 开始 ---"
    For i = 0 To UBound(arr)
        Debug.Print i & " => " & arr(i)
    Next
    Debug.Print "--- 结束 ---"
End Function

' 打印外部一维数组
Function printArrOther(arr As Variant)
    Debug.Print "--- 开始 ---"
    For i = 0 To UBound(arr)
        Debug.Print i & " => " & arr(i)
    Next
    Debug.Print "--- 结束 ---"
End Function


