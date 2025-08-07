# .NET PowerPoint Library Performance Benchmarks

**Overview**  
The Syncfusion® .NET PowerPoint library (Presentation) enables seamless integration for working with PowerPoint files, offering robust features for handling presentations in various formats. This performance benchmark report demonstrates the speed and efficiency of key functionalities, emphasizing how our library excels in real-world scenarios.

**Environment**  
The following system configurations were used for benchmarking:
- Operating System: Windows 10  
- Processor: 11th Gen Intel(R) Core(TM)  
- RAM: 16GB  
- .NET Version: .NET 8.0  
- Syncfusion® Version: Syncfusion.Presentation.Net.Core v30.1.37  

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
    <tr><td>2</td><td>PowerPoint-2.pptx</td><td>0.01</td></tr>
    <tr><td>50</td><td>PowerPoint-50.pptx</td><td>0.02</td></tr>
    <tr><td>100</td><td>PowerPoint-100.pptx</td><td>0.1</td></tr>
    <tr><td>500</td><td>PowerPoint-500.pptx</td><td>2.6</td></tr>
  </tbody>
</table>

You can find the sample used for this performance evaluation on [GitHub](https://github.com/SyncfusionExamples/PowerPoint-Examples/tree/master/Performance-metrices/Open-and-save/).

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
    <tr><td>2</td><td>PowerPoint-2.pptx</td><td>0.01</td></tr>
    <tr><td>50</td><td>PowerPoint-50.pptx</td><td>0.02</td></tr>
    <tr><td>100</td><td>PowerPoint-100.pptx</td><td>0.06</td></tr>
    <tr><td>500</td><td>PowerPoint-500.pptx</td><td>0.5</td></tr>
  </tbody>
</table>

You can find the sample used for this performance evaluation on [GitHub](https://github.com/SyncfusionExamples/PowerPoint-Examples/tree/master/Performance-metrices/Clone-and-merge-slides/).

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
    <tr><td>2</td><td>PowerPoint-2.pptx</td><td>0.03</td></tr>
    <tr><td>50</td><td>PowerPoint-50.pptx</td><td>0.1</td></tr>
    <tr><td>100</td><td>PowerPoint-100.pptx</td><td>1.2</td></tr>
    <tr><td>500</td><td>PowerPoint-500.pptx</td><td>16</td></tr>
  </tbody>
</table>

You can find the sample used for this performance evaluation on [GitHub](https://github.com/SyncfusionExamples/PowerPoint-Examples/tree/master/Performance-metrices/PPTX-to-PDF/).

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
    <tr><td>2</td><td>PowerPoint-2.pptx</td><td>0.05</td></tr>
    <tr><td>50</td><td>PowerPoint-50.pptx</td><td>0.4</td></tr>
    <tr><td>100</td><td>PowerPoint-100.pptx</td><td>2.8</td></tr>
    <tr><td>500</td><td>PowerPoint-500.pptx</td><td>32</td></tr>
  </tbody>
</table>

You can find the sample used for this performance evaluation on [GitHub](https://github.com/SyncfusionExamples/PowerPoint-Examples/tree/master/Performance-metrices/PPTX-to-Image/).
