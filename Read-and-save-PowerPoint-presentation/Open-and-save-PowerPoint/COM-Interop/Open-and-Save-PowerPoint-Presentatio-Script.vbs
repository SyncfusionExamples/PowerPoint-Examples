Option Explicit
On Error Resume Next

' --- Change this input and output path as you want ---
Dim inputPath
inputPath = "D:\Template.pptx"

Dim outputPath
outputPath = "D:\Output.pptx"

' Create the COM object exposed by your .NET class
' ProgID must match the [ProgId] attribute in your C# class
Dim obj
Set obj = CreateObject("Open_and_Save_PowerPoint_Presentation.Class1")

If Err.Number <> 0 Then
    MsgBox "Failed to create COM object 'Open_and_Save_PowerPoint_Presentation.Class1'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "COM Error"
    WScript.Quit 1
End If

' Call the method exposed by your class
obj.OpenAndSavePowerPointPresentation inputPath, outputPath

If Err.Number <> 0 Then
    MsgBox "Method call failed (OpenAndSavePowerPointPresentation)." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Invoke Error"
    WScript.Quit 2
End If

' Release the COM object
Set obj = Nothing

MsgBox "Success!" & vbCrLf & "Document created at:" & vbCrLf & outputPath, vbInformation, "Done"
