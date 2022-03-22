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

' ��������
Property Get data() As Variant
    data = arr
End Property

' ��ʼ������
Function init(row As Integer)
    ReDim arr(row, 0)
End Function

'�����ַ�
Function addStr(row As Integer, str As String)
    Dim col As Integer
    col = UBound(arr, 2) '�г��ȣ�ÿһ�ж���һ����,ֻҪ��һ�м�1�������ж�����ż�1
    
    '����ָ�����У�����ڵ�һ���յ�λ��
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

'��������
Function addInt(row As Integer, num As Integer)
    Dim col As Integer
    col = UBound(arr, 2) '�г��ȣ�ÿһ�ж���һ����,ֻҪ��һ�м�1�������ж�����ż�1
    
    '����ָ�����У�����ڵ�һ���յ�λ��
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

' ��ӡ��ά����
Function printArr()
    Debug.Print "--- ��ʼ ---"
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
    Debug.Print "--- ���� ---"
End Function

' ��ӡ�ⲿ��ά����
Function printArrOther(arr As Variant)
    Debug.Print "--- ��ʼ ---"
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
    Debug.Print "--- ���� ---"
End Function