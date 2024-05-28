VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get GetFormatIndex(ByVal Format As String, ByRef formatArray As Variant) As Long
    Dim i As Long
    ' Find the index of the format in the array
    For i = LBound(formatArray) To (UBound(formatArray) - 1)
        If Format = formatArray(i) Then
            GetFormatIndex = i + 1
            Exit Property
        End If
    Next i
    
    ' If format is not found, return 0
    GetFormatIndex = 0
End Property



Public Property Get GetFormatColorIndex(ByVal Format As Variant, ByRef formatArray As Variant) As Long
    Dim i As Long
    ' Find the index of the format in the array
    For i = LBound(formatArray) To (UBound(formatArray) - 1)
        If AreArraysEqual(Format, formatArray(i)) Then
            GetFormatColorIndex = i + 1
            Exit Property
        End If
    Next i
    
    ' If format is not found, return 0
    GetFormatColorIndex = 0
End Property






Function AreArraysEqual(arr1 As Variant, arr2 As Variant) As Boolean
    Dim i As Long
    
    ' Compare the lengths of the arrays
    If UBound(arr1) - LBound(arr1) <> UBound(arr2) - LBound(arr2) Then
        AreArraysEqual = False
        Exit Function
    End If
    
    ' Compare each element of the arrays
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            AreArraysEqual = False
            Exit Function
        End If
    Next i
    
    ' All elements are equal
    AreArraysEqual = True
End Function




Function GetFontColorRGBArray(targetCell As Range) As Variant
    Dim fontColor As Long
    Dim red As Long, green As Long, blue As Long
    Dim rgbArray(0 To 2) As Long
    
    ' Get the font color of the cell
    fontColor = targetCell.Font.Color
    
    ' Decompose the RGB color into red, green, and blue components
    red = fontColor Mod 256
    green = (fontColor \ 256) Mod 256
    blue = (fontColor \ 256 \ 256) Mod 256
    
    ' Store the RGB values in an array
    rgbArray(0) = red
    rgbArray(1) = green
    rgbArray(2) = blue
    
    GetFontColorRGBArray = rgbArray
End Function

