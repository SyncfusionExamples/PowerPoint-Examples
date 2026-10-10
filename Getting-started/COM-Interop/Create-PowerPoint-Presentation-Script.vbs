Option Explicit
On Error Resume Next

' --- Change this output path as you want ---
Dim outputPath
outputPath = "D:\Output.pptx"

' Create the COM object exposed by your .NET class
' ProgID must match the [ProgId] attribute in your C# class
Dim obj
Set obj = CreateObject("Create_PowerPoint_Presentation.Class1")

If Err.Number <> 0 Then
    MsgBox "Failed to create COM object 'Create_PowerPoint_Presentation.Class1'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "COM Error"
    WScript.Quit 1
End If

' Call the method exposed by your class
obj.CreatePowerPointPresentation outputPath

If Err.Number <> 0 Then
    MsgBox "Method call failed (CreatePowerPointPresentation)." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Invoke Error"
    WScript.Quit 2
End If

' Release the COM object
Set obj = Nothing

MsgBox "Success!" & vbCrLf & "Document created at:" & vbCrLf & outputPath, vbInformation, "Done"
