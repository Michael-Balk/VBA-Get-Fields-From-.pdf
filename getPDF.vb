Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices.JavaScript

Public Class getPDF
    Public Function getFormData(formLocation As String) As Array

        Dim AcrobatApp As Object
        Dim thePDF As Object
        Dim javascriptObj As Object

        AcrobatApp = CreateObject("AcroExch.App")
        thePDF = CreateObject("AcroExch.PDDoc")

        thePDF.Open(formLocation)

        javascriptObj = thePDF.GetJSObject

        Dim num As Integer
        num = javascriptObj.numFields

        MsgBox("The Number of Fields is: " + CStr(num))

        Dim result(num - 1) As Object

        For i = 0 To num - 1
            result(i) = javascriptObj.getField(javascriptObj.getNthFieldName(i)).Value
        Next i

        thePDF.Close
        AcrobatApp.Exit
        AcrobatApp = Nothing
        thePDF = Nothing


        getFormData = result

    End Function
End Class