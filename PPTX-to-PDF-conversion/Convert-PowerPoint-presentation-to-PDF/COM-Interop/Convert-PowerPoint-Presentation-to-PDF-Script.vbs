Option Explicit
On Error Resume Next

' --- Change this input and output path as you want ---
Dim inputPath
inputPath = "D:\Template.pptx"

Dim outputPath
outputPath = "D:\Output.pdf"

' Create the COM object exposed by your .NET class
' ProgID must match the [ProgId] attribute in your C# class
Dim obj
Set obj = CreateObject("Convert_PowerPoint_Presentation_to_PDF.Class1")

If Err.Number <> 0 Then
    MsgBox "Failed to create COM object 'Convert_PowerPoint_Presentation_to_PDF.Class1'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "COM Error"
    WScript.Quit 1
End If

' Call the method exposed by your class
obj.ConvertPowerPointPresentationtoPDF inputPath, outputPath

If Err.Number <> 0 Then
    MsgBox "Method call failed (ConvertPowerPointPresentationtoPDF)." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Invoke Error"
    WScript.Quit 2
End If

' Release the COM object
Set obj = Nothing

MsgBox "Success!" & vbCrLf & "Document created at:" & vbCrLf & outputPath, vbInformation, "Done"
