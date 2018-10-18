<div align="right">
<a href="/">Blog</a> &nbsp;
<a href="/about">About</a>
<hr>
</div>

# Apply Cell Borders on Excel Worksheet Cells in C++

Microsoft Excel offers various types of built-in borders that can be applied to a cell or range of cells inside the worksheet. Excel cell border consists of two components i.e. line style and color. Since, there are number of line styles available in Microsoft Excel, user can also create custom borders. [Aspose.Cells for C++](https://products.aspose.com/cells/cpp) API can be used to create any type of border, be it built-in border or custom border with ease. Besides, it can be used to create, edit and manipulate Excel spreadsheets almost in any platform without any need to install Microsoft Excel or without using any sort of Microsoft Office automation.

## Article Description

The purpose of this article is to explain how developers can apply cell borders on Excel worksheet cells in C++.

## Supported Platforms

Aspose.Cells API supports various platforms including C++, Java, .NET, Android, JavaScript, PHP etc. Besides, [Aspose.Cells is available in Cloud as RESTful APIs](https://products.aspose.cloud/cells).

# Types of Built-in Borders

Following are the types of built-in borders in Microsoft Excel that you can apply to cell.

* Top
* Bottom
* Left
* Right
* Diagonal Up
* Diagonal Down

# Cell Border Types To Create Custom Borders

By changing line styles, you can create custom borders. Some of these line styles are mentioned below. Please also see the snapshot given below that displays line styles in Microsoft Excel Border GUI.

* None
* Medium
* Hair
* Thick
* Thin
* Dashed
* Dotted
* Dash Dotted

>_**Caption:** Border Line Styles in Microsoft Excel represented by CellBorderType in Aspose.Cells API._

![Border Line Styles in Microsoft Excel represented by CellBorderType in Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Apply%20Cell%20Borders%20on%20Excel%20Worksheet%20Cells/Border-Line-Styles-Microsoft-Excel-Border-Aspose.Cells.png "Border Line Styles in Microsoft Excel represented by CellBorderType in Aspose.Cells API.")

# Set Border of the Cell using Aspose.Cells API

All of the border types can be accessed using the BorderType enumeration. Its values are as follows

* Aspose::Cells::BorderType_TopBorder
* Aspose::Cells::BorderType_BottomBorder
* Aspose::Cells::BorderType_LeftBorder
* Aspose::Cells::BorderType_RightBorder
* Aspose::Cells::BorderType_DiagonalUp
* Aspose::Cells::BorderType_DiagonalDown

The following code accesses the top border of the cell and sets its line style and color. Similarly, you can work with any border using the BorderType enumeration.

```javascript
// Set Top Border of Cell.
// --------------------------

// Access cell object.
intrusive_ptr<Aspose::Cells::ICell> cell = ws->GetICells()->GetObjectByIndex(new String("B3"));

// Access cell style.
intrusive_ptr<Aspose::Cells::IStyle> style = cell->GetIStyle();

// Access top border.
intrusive_ptr<Aspose::Cells::IBorder> topBorder = style->GetIBorders()->GetObjectByIndex(Aspose::Cells::BorderType_TopBorder);

// Set the line style of the top border.
topBorder->SetLineStyle(Aspose::Cells::CellBorderType_Thick);

// Set the color of the top border.
topBorder->SetColor(Aspose::Cells::System::Drawing::Color::GetRed());

// Set the cell style.
cell->SetIStyle(style);
```
