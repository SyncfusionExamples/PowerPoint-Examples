# .NET PowerPoint Library Performance Benchmarks

**Overview**  
The Syncfusion® .NET PowerPoint library (Presentation) enables seamless integration for working with PowerPoint files, offering robust features for handling presentations in various formats. This performance benchmark report demonstrates the speed and efficiency of key functionalities, emphasizing how our library excels in real-world scenarios.

**Environment**  
The following system configurations were used for benchmarking:
- Operating System: Windows 10  
- Processor: 11th Gen Intel(R) Core(TM)  
- RAM: 16GB  
- .NET Version: .NET 8.0  
- Syncfusion® Version: [Syncfusion.PresentationRenderer.Net.Core v31.1.17](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core/31.1.17) ,   [Syncfusion.Presentation.Net.Core v31.1.17](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core/31.1.17)

**Open and Save Presentation**

<table>
  <thead>
    <tr>
      <th>Slides</th>
      <th>Input Presentation File</th>
      <th>Syncfusion® Time (sec)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>2</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-2.pptx">PowerPoint-2.pptx</a></td><td>0.01</td></tr>
    <tr><td>50</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-50.pptx</a></td><td>0.02</td></tr>
    <tr><td>100</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-100.pptx">PowerPoint-2.pptx</a></td><td>0.1</td></tr>
    <tr><td>500</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-500.pptx">PowerPoint-500.pptx</a></td><td>2.6</td></tr>
  </tbody>
</table>

**Clone and Merge Slides**

<table>
  <thead>
    <tr>
      <th>Slides</th>
      <th>Input Presentation File</th>
      <th>Syncfusion® Time (sec)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>2</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-2.pptx">PowerPoint-2.pptx</a></td><td>0.01</td></tr>
    <tr><td>50</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-50.pptx</a></td><td>0.02</td></tr>
    <tr><td>100</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-100.pptx</a></td><td>0.06</td></tr>
    <tr><td>500</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-500.pptx">PowerPoint-500.pptx</a></td><td>0.5</td></tr>
  </tbody>
</table>

**PowerPoint to PDF Conversion**

<table>
  <thead>
    <tr>
      <th>Slides</th>
      <th>Input Presentation File</th>
      <th>Syncfusion® Time (sec)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>2</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-2.pptx">PowerPoint-2.pptx</a></td><td>0.03</td></tr>
    <tr><td>50</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-50.pptx</a></td><td>0.1</td></tr>
    <tr><td>100</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-100.pptx</a></td><td>1.2</td></tr>
    <tr><td>500</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-500.pptx">PowerPoint-500.pptx</a></td><td>16</td></tr>
  </tbody>
</table>

**PowerPoint to Image Conversion**

<table>
  <thead>
    <tr>
      <th>Slides</th>
      <th>Input Presentation File</th>
      <th>Syncfusion® Time (sec)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>2</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-2.pptx">PowerPoint-2.pptx</a></td><td>0.05</td></tr>
    <tr><td>50</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-50.pptx</a></td><td>0.4</td></tr>
    <tr><td>100</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-50.pptx">PowerPoint-100.pptx</a></td><td>2.8</td></tr>
    <tr><td>500</td><td><a href="https://github.com/SyncfusionExamples/PowerPoint-Examples/blob/master/Performance-metrices/PPTX-to-Image/.NET/Convert-PowerPoint-slide-to-Image/Data/PowerPoint-500.pptx">PowerPoint-500.pptx</a></td><td>32</td></tr>
  </tbody>
</table>
