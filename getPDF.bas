Attribute VB_Name = "Module1"
Sub runFunction()

    Dim Result() As Variant

    Result = getFormData("C:\.....\Sample Form.pdf")

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
        
        Dim fieldNames() As Variant
        fieldNames = Array("fieldOne", "fieldTwo", "fieldThree") ''these are the known field names from the pdf
        
        For i = 0 To UBound(fieldNames) 'num - 1
        If IsNull(javascriptObj.getField(fieldNames(i)).Value) = False Then
            ''Result(i) = javasript.getField(javascriptObj.getNthFieldName(i)).Value  ''getNthFieldName works in .VB, haven't got it working here
            Result(i) = javascriptObj.getField(fieldNames(i)).Value     ''using known field names array instead
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



