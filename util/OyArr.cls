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

' ��������
Property Get data() As Variant
    data = arr
End Property

' ��ʼ������
Function init()
    ReDim arr(0)
End Function

'�����ַ�
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

'��������
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

' ��ӡһά����
Function printArr()
    Debug.Print "--- ��ʼ ---"
    For i = 0 To UBound(arr)
        Debug.Print i & " => " & arr(i)
    Next
    Debug.Print "--- ���� ---"
End Function

' ��ӡ�ⲿһά����
Function printArrOther(arr As Variant)
    Debug.Print "--- ��ʼ ---"
    For i = 0 To UBound(arr)
        Debug.Print i & " => " & arr(i)
    Next
    Debug.Print "--- ���� ---"
End Function

