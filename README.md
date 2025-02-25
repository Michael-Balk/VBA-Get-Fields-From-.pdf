# VBA-Get-Fields-From-.pdf
VBA to retrieve Form Fields from a fillable .pdf

Also included in this repository is a VB class with the same functionality for .net projects.

```VBA
Sub runFunction()

    Dim Result() As Variant

    Result = getFormData("C:\Users\mbalk\OneDrive - State of Kansas, OITS\Desktop\Sample Form.pdf")

    For i = 0 To UBound(Result)
        MsgBox CStr(Result(i))
    Next i

End Sub

Function getFormData(formLocation As String)

        Dim AcrobatApp As Object
        Dim thePDF As Object
        Dim javascriptObj As Object
        Dim Result() As Variant
        Set AcrobatApp = CreateObject("AcroExch.App")
        Set thePDF = CreateObject("AcroExch.PDDoc")

        thePDF.Open (formLocation)

        Set javascriptObj = thePDF.GetJSObject

        Dim num As Integer
        num = 0
        num = javascriptObj.numFields

        MsgBox ("The Number of Fields is: " + CStr(num))

        ReDim Result(num - 1) As Variant
        
        For i = 0 To num - 1
        If IsNull(javascriptObj.getField(javascriptObj.getNthFieldName(i)).Value) = False Then
            Result(i) = javascriptObj.getField(javascriptObj.getNthFieldName(i)).Value
        Else
            Result(i) = Null
        End If

        Next i

        
        thePDF.Close
        AcrobatApp.Exit
        Set AcrobatApp = Nothing
        Set thePDF = Nothing


        getFormData = Result

End Function
