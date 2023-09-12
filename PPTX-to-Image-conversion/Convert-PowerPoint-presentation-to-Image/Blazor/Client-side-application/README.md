Convert PowerPoint to images in Blazor WebAssembly (WASM)
---------------------------------------------------------

This example illustrates how to convert PowerPoint to images in Blazor WebAssembly (WASM).

Steps to convert PowerPoint to images in Blazor WebAssembly (WASM)
------------------------------------------------------------------

1. Create a new C# Blazor WebAssembly App in Visual Studio.
2. Install the [Syncfusion.PresentationRenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core) NuGet package as a reference to your Blazor application from [NuGet.org](https://www.nuget.org/).  
3. Install the [SkiaSharp.NativeAssets.WebAssembly](https://www.nuget.org/packages/SkiaSharp.NativeAssets.WebAssembly) NuGet package as a reference to your Blazor application from [NuGet.org](https://www.nuget.org/).  
4. Add the following ItemGroup tag in the [Blazor WASM csproj](https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/PPTX-to-Image-conversion/Convert-PowerPoint-presentation-to-Image/Blazor/Client-side-application/Convert-PPTX-to-Image/Convert-PPTX-to-Image.csproj) file.

```xml
<ItemGroup>
    <NativeFileReference Include="$(SkiaSharpStaticLibraryPath)\2.0.23\*.a" />
</ItemGroup>
```

N> Install this `wasm-tools` and `wasm-tools-net6` by using `dotnet workload install wasm-tools` and `dotnet workload install wasm-tools-net6` commands in your command prompt respectively, while facing issues related to skiasharp, during runtime.

5. Enable the following property in the [Blazor WASM csproj](https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/PPTX-to-Image-conversion/Convert-PowerPoint-presentation-to-Image/Blazor/Client-side-application/Convert-PPTX-to-Image/Convert-PPTX-to-Image.csproj) file.

```xml
<PropertyGroup>
    <WasmNativeStrip>true</WasmNativeStrip>
</PropertyGroup>
```

6. Create a [razor](https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/PPTX-to-Image-conversion/Convert-PowerPoint-presentation-to-Image/Blazor/Client-side-application/Convert-PPTX-to-Image/Pages/DocIO.razor) file named as Presentation under the **Pages** folder and add the namespaces in the file.
7. Add the code to create a button.
8. Create a new async method with the name PPTXToImage and include the code sample to convert a PowerPoint to images.
9. Create a [class](https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/PPTX-to-Image-conversion/Convert-PowerPoint-presentation-to-Image/Blazor/Client-side-application/Convert-PPTX-to-Image/FileUtils.cs) file with FileUtils name and add the code to invoke the JavaScript action to download the file in the browser.
10. Add the JavaScript function in the [Index.html](https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/PPTX-to-Image-conversion/Convert-PowerPoint-presentation-to-Image/Blazor/Client-side-application/Convert-PPTX-to-Image/wwwroot/index.html) file present under the **wwwroot** folder.
11. Add the code sample in the [razor](https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/PPTX-to-Image-conversion/Convert-PowerPoint-presentation-to-Image/Blazor/Client-side-application/Convert-PPTX-to-Image/Shared/NavMenu.razor) file of the Navigation menu in the **Shared** folder.
12. Rebuild the solution.
13. Run the application.
