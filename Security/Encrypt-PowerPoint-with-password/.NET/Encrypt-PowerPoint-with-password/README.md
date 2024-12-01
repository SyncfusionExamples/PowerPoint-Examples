# Encrypt PowerPoint Presentation using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **encrypt a PowerPoint Presentation** using C#.

## Steps to encrypt a PowerPoint Presentation programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
```

Step 4: Add the following code snippet in Program.cs file to encrypt the PowerPoint Presentation.

```csharp
//Open an existing presentation.
using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read);
using IPresentation presentation = Presentation.Open(inputStream);
//Encrypt the presentation with a password.
presentation.Encrypt("syncfusion");
using FileStream outputStream = new (Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
presentation.Save(outputStream);
```

More information about encrypting presentations can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/security) section.