VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OyArr2"
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
Function init(row As Integer)
    ReDim arr(row, 0)
End Function

'添加字符
Function addStr(row As Integer, str As String)
    Dim col As Integer
    col = UBound(arr, 2) '列长度，每一行都是一样的,只要有一行加1，其他行都会跟着加1
    
    '遍历指定的行，填充在第一个空的位置
    For i = 0 To col
        If IsEmpty(arr(row, i)) Then
            arr(row, i) = str
            Exit For
        ElseIf i = col Then
            ReDim Preserve arr(2, i + 1)
            arr(row, i + 1) = str
            Exit For
        End If
    Next
End Function

'添加数字
Function addInt(row As Integer, num As Integer)
    Dim col As Integer
    col = UBound(arr, 2) '列长度，每一行都是一样的,只要有一行加1，其他行都会跟着加1
    
    '遍历指定的行，填充在第一个空的位置
    For i = 0 To col
        If IsEmpty(arr(row, i)) Then
            arr(row, i) = num
            Exit For
        ElseIf i = col Then
            ReDim Preserve arr(2, i + 1)
            arr(row, i + 1) = num
            Exit For
        End If
    Next
End Function

' 打印二维数组
Function printArr()
    Debug.Print "--- 开始 ---"
    For i = 0 To UBound(arr, 1)
        Dim str As String
        str = ""
        For j = 0 To UBound(arr, 2)
            If j = 0 Then
                str = str & arr(i, j)
            Else
                str = str & "," & arr(i, j)
            End If
        Next
        Debug.Print i & " => " & str
    Next
    Debug.Print "--- 结束 ---"
End Function

' 打印外部二维数组
Function printArrOther(arr As Variant)
    Debug.Print "--- 开始 ---"
    For i = 0 To UBound(arr, 1)
        Dim str As String
        str = ""
        For j = 0 To UBound(arr, 2)
            If j = 0 Then
                str = str & arr(i, j)
            Else
                str = str & "," & arr(i, j)
            End If
        Next
        Debug.Print i & " => " & str
    Next
    Debug.Print "--- 结束 ---"
End Function
