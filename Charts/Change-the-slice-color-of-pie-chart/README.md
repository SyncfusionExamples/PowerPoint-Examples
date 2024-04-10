# Change the slice color of Pie chart 
To change the slice color of pie chart in PowerPoint Presentation using .NET PowerPoint library, please refer the below code example

 

```
//Change the color of each slice
serie.DataPoints[0].DataFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Cornsilk;
serie.DataPoints[1].DataFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Violet;
serie.DataPoints[2].DataFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Pink;
```