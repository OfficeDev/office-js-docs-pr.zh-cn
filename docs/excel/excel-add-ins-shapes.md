---
title: 使用 JavaScript API Excel形状
description: 了解如何Excel形状定义为位于绘图层上的任何Excel。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 533a9cf9689bcaa5cd43635da836730a2af6ab61
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938750"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>使用 JavaScript API Excel形状

Excel形状定义为位于绘图层上的任何Excel。 这意味着单元格之外任何内容都是形状。 本文介绍如何将几何形状、线条和图像与 [Shape](/javascript/api/excel/excel.shape) 和 [ShapeCollection](/javascript/api/excel/excel.shapecollection) API 结合使用。 [图](/javascript/api/excel/excel.chart)在其自己的文章使用[JavaScript API](excel-add-ins-charts.md)处理图表Excel介绍。

下图显示了形成温度计的形状。
![作为温度计的形状Excel的图像。](../images/excel-shapes.png)

## <a name="create-shapes"></a>创建形状

形状通过工作表的形状集合创建并存储在 `Worksheet.shapes` () 。 `ShapeCollection` 为此 `.add*` ，有几个方法。 所有形状在添加到集合时都有为它们生成的名称和 ID。 它们分别为 `name` 和 `id` 属性。 `name` 可通过外接程序设置 ，以使用 方法轻松 `ShapeCollection.getItem(name)` 检索。

以下类型的形状是使用关联方法添加的。

| 形状 | Add 方法 | 签名 |
|-------|------------|-----------|
| 几何形状 | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 图像 (JPEG 或 PNG)  | [addImage](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_) | `addImage(base64ImageString: string): Excel.Shape` |
| 折线图 | [addLine](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#addSvg_xml_) | `addSvg(xml: string): Excel.Shape` |
| 文本框 | [addTextBox](/javascript/api/excel/excel.shapecollection#addTextBox_text_) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>几何形状

使用 创建几何形状 `ShapeCollection.addGeometricShape` 。 该方法采用 [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) 枚举作为参数。

下面的代码示例创建一个名为 **"Square"** 的 150x150 像素矩形，该矩形从工作表的顶部和左侧放置 100 像素。

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="images"></a>图像

JPEG、PNG 和 SVG 图像可以作为形状插入到工作表中。 `ShapeCollection.addImage`方法采用 base64 编码的字符串作为参数。 这是字符串形式的 JPEG 或 PNG 图像。 `ShapeCollection.addSvg` 即使此参数是定义图形的 XML，也采用字符串。

下面的代码示例演示 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 作为字符串加载的图像文件。 该字符串具有元数据"base64"，该元数据在形状创建之前被删除。

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

使用 创建行 `ShapeCollection.addLine` 。 该方法需要线条起点和终点的左边距和上边距。 它还需要 [一个 ConnectorType](/javascript/api/excel/excel.connectortype) 枚举来指定终结点之间的直线连接。 下面的代码示例在工作表上创建一条直线。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

线条可以连接到其他 Shape 对象。 和 `connectBeginShape` `connectEndShape` 方法将线条的开头和结尾附加到指定连接点处的形状。 这些点的位置因形状而异，但 可用于确保您的外接程序不会连接到外部 `Shape.connectionSiteCount` 的点。 线条使用 和 方法断开与任何附加 `disconnectBeginShape` `disconnectEndShape` 形状的连接。

下面的代码示例将 **"MyLine"** 行连接到名为 **"LeftShape"** 和 **"RightShape"的两个形状**。

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-and-resize-shapes"></a>移动形状并调整形状大小

形状位于工作表的顶部。 它们的位置由 和 `left` `top` 属性定义。 它们充当工作表各自边缘的边距，[0， 0] 为左上角。 可以使用 和 方法直接设置它们，也可以从当前位置 `incrementLeft` `incrementTop` 进行调整。 此外，还通过此方式建立形状相对于默认位置的旋转量，该属性为绝对量，而方法用于调整现有 `rotation` `incrementRotation` 旋转。

形状相对于其他形状的深度由 属性 `zorderPosition` 定义。 这是使用 方法 `setZOrder` 设置的，该方法采用 [ShapeZOrder](/javascript/api/excel/excel.shapezorder)。 `setZOrder` 调整当前形状相对于其他形状的排序。

加载项具有多个选项，用于更改形状的高度和宽度。 如果设置 `height` 或 `width` 属性，则更改指定维度而不更改其他维度。 和 `scaleHeight` `scaleWidth` 根据所提供的 [ShapeScaleType](/javascript/api/excel/excel.shapescaletype) 属性的值调整形状相对于当前大小 (或原始尺寸) 。 可选的 [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) 参数指定形状从何处缩放 (左上角、中间或右下角缩放) 。 如果 `lockAspectRatio` 该属性为 **true，** 则缩放方法通过同时调整其他尺寸来保持形状的当前纵横比。

> [!NOTE]
> 对 和 `height` `width` 属性的直接更改仅影响该属性，而 `lockAspectRatio` 不管属性值如何。

下面的代码示例显示一个形状，其缩放比例为其原始大小的 1.25 倍，并旋转 30 度。

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="text-in-shapes"></a>形状中的文本

几何形状可以包含文本。 形状具有 `textFrame` [TextFrame 类型的属性](/javascript/api/excel/excel.textframe)。 对象 `TextFrame` 管理文本显示选项 (如边距和文本溢出) 。 `TextFrame.textRange` 是文本内容和字体设置的 [TextRange](/javascript/api/excel/excel.textrange) 对象。

下面的代码示例创建一个名为"Wave"的几何形状，其文本为"Shape Text"。 它还调整形状和文本颜色，以及将文本的水平对齐方式设置到中心。

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;
    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");
    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    return context.sync();
}).catch(errorHandlerFunction);
```

`addTextBox`方法创建 `ShapeCollection` 具有 `GeometricShape` 白色 `Rectangle` 背景和黑色文本的 类型。 这与"插入"选项卡 **Excel"文本框**"按钮 **所创建** 的内容相同。 `addTextBox`采用字符串参数来设置 的文本 `TextRange` 。

下面的代码示例演示了如何创建文本为"Hello！"的文本框。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="shape-groups"></a>形状组

形状可以组合在一起。 这允许用户将它们视为用于定位、调整大小和其他相关任务的单个实体。 [ShapeGroup](/javascript/api/excel/excel.shapegroup)是 的一种类型，因此加载项将组 `Shape` 视为单个形状。

下面的代码示例显示了组合在一起的三个形状。 后续代码示例显示形状组向右移动 50 像素。

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var square = shapes.getItem("Square");
    var pentagon = shapes.getItem("Pentagon");
    var octagon = shapes.getItem("Octagon");

    var shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    return context.sync();
}).catch(errorHandlerFunction);

// This sample moves the previously created shape group to the right by 50 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shapeGroup = sheet.shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> 组内的单个形状通过属性引用，该属性的类型 `ShapeGroup.shapes` 为 [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection)。 分组后，无法再通过工作表的形状集合访问它们。 例如，如果工作表有三个形状，并且它们全部组合在一起，则工作表 `shapes.getCount` 的方法将返回计数 1。

## <a name="export-shapes-as-images"></a>将形状导出为图像

任何 `Shape` 对象都可以转换为图像。 [Shape.getAsImage](/javascript/api/excel/excel.shape#getAsImage_format_) 返回 base64 编码的字符串。 图像的格式指定为传递给 的 [PictureFormat](/javascript/api/excel/excel.pictureformat) 枚举 `getAsImage` 。

```js
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shape = sheet.shapes.getItem("Image");
    var stringResult = shape.getAsImage(Excel.PictureFormat.png);

    return context.sync().then(function () {
        console.log(stringResult.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

## <a name="delete-shapes"></a>删除形状

使用对象的方法从工作表 `Shape` 中删除 `delete` 形状。 无需其他元数据。

下面的代码示例从 **MyWorksheet 中删除所有形状**。

```js
// This deletes all the shapes from "MyWorksheet".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
            shape.delete()
        });
        return context.sync();
    }).catch(errorHandlerFunction);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)
- [使用 Excel JavaScript API 处理图表](excel-add-ins-charts.md)
