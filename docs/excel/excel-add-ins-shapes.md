---
title: 使用 Excel JavaScript API 处理形状
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: e4d01c387fff01d68cb26369240a1e06e723a54c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926647"
---
# <a name="work-with-shapes-using-the-excel-javascript-api-preview"></a>使用 Excel JavaScript API 处理形状 (预览)

> [!NOTE]
> 本文中讨论的 api 目前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

excel 将形状定义为位于 Excel 绘图层的任何对象。 这意味着单元格之外的任何内容都是形状。 本文介绍如何将几何形状、线条和图像与 [Shape]/javascript/api/excel/excel.shape) 和[ShapeCollection](/javascript/api/excel/excel.shapecollection) api 结合使用。 [图表](/javascript/api/excel/excel.chart)在其自己的文章中进行了介绍, [使用 Excel JavaScript API 处理图表]] (Excel 外接程序-charts.md))。

## <a name="create-shapes"></a>创建形状

形状是通过工作表的形状集合 (`Worksheet.shapes`) 创建和存储的。 `ShapeCollection`有几`.add*`种方法可以实现此目的。 在将所有形状添加到集合中时, 都会为它们生成名称和 id。 它们分别是`name`和`id`属性。 `name`可通过外接程序进行设置以便使用`ShapeCollection.getItem(name)`方法轻松检索。

使用关联的方法添加以下类型的形状:

| Shape | Add 方法 | 签名 |
|-------|------------|-----------|
| 几何形状 | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 图像 (JPEG 或 PNG) | [addImage](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| 文本框 | [addTextBox](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>几何形状

将使用`ShapeCollection.addGeometricShape`创建一个几何形状。 该方法采用[GeometricShapeType](/javascript/api/excel/excel.geometricshapetype)枚举作为参数。

下面的代码示例创建一个名为 **"** 150x150" 的像素矩形, 该矩形在工作表的顶部和左侧位置为100像素。

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

JPEG、PNG 和 SVG 图像可以作为形状插入到工作表中。 该`ShapeCollection.addImage`方法采用 base64 编码的字符串作为参数。 这是字符串形式的 JPEG 或 PNG 图像。 `ShapeCollection.addSvg`也采用字符串, 但此参数是用于定义图形的 XML。

下面的代码示例显示了[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader)作为字符串加载的图像文件。 字符串中包含元数据 "base64", 在创建形状之前将其删除。

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = event.target.result.indexOf("base64,");
        var myBase64 = event.target.result.substr(startIndex + 7);
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

将使用`ShapeCollection.addLine`创建的行。 该方法需要线条的起始点和结束点的左边距和上边距。 它还采用[ConnectorType](/javascript/api/excel/excel.connectortype)枚举来指定行在终结点之间的 contorts 方式。 下面的代码示例在工作表上创建一条直线。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

可以将线条连接到其他 Shape 对象。 `connectBeginShape`和`connectEndShape`方法将行的开头和结尾附加到指定连接点处的形状。 这些点的位置因形状而异, 但`Shape.connectionSiteCount`可用于确保外接程序不会连接到超出边界的点。 使用`disconnectBeginShape`和`disconnectEndShape`方法将线条与任何附加的形状断开连接。

下面的代码示例将 **"MyLine"** 行连接到名为 **"LeftShape"** 和 **"RightShape"** 的两个形状。

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

## <a name="move-and-resize-shapes"></a>移动形状并调整其大小

形状位于工作表的顶部。 它们的位置由`left`和`top`属性定义。 这些操作充当工作表各自边缘的边距, [0, 0] 为左上角。 这些值既可以直接设置`incrementLeft` , 也可以使用 and `incrementTop`方法从当前位置进行调整。 从默认位置旋转的形状也是通过这种方式建立的, `rotation`属性是绝对量和调整现有旋转的`incrementRotation`方法。

相对于其他形状的形状的深度由该`zorderPosition`属性定义。 这是使用`setZOrder`方法进行设置的, 该方法采用[ShapeZOrder](/javascript/api/excel/excel.shapezorder)。 `setZOrder`调整当前形状相对于其他形状的排序。

您的外接程序有几个用于更改形状的高度和宽度的选项。 设置`height`或`width`属性将更改指定的维度, 而不更改其他维度。 和调整相对于当前或原始尺寸的形状各自的尺寸 (基于提供的 ShapeScaleType 的值)。 [](/javascript/api/excel/excel.shapescaletype) `scaleHeight` `scaleWidth` 可选的[ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom)参数指定形状的缩放位置 (左上角、中间或右下角)。 如果该`lockAspectRatio`属性为**true**, 则缩放方法还会通过调整其他尺寸来保持形状的当前纵横比。

> [!NOTE]
> 对`height`和`width`属性的直接更改仅影响该属性, 而不考虑`lockAspectRatio`属性的值。

下面的代码示例显示了缩放到其原始大小为1.25 倍且旋转30度的形状。

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

几何形状可以包含文本。 形状具有类型`textFrame`为[TextFrame](/javascript/api/excel/excel.textframe)的属性。 `TextFrame`对象管理文本显示选项 (如边距和文本溢出)。 `TextFrame.textRange`是一个带有 "文本内容" 和 "字体" 设置的[TextRange](/javascript/api/excel/excel.textrange)对象。

下面的代码示例创建一个名为 "Wave" 的几何形状, 其中包含文本 "shape text"。 它还调整形状和文本颜色, 并将文本的水平对齐方式设置为居中。

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

创建具有白色背景和`GeometricShape`黑色文本`Rectangle`的类型的`addTextBox`方法。 `ShapeCollection` 这与 Excel 的 "**插入**" 选项卡上的 "**文本框**" 按钮所创建的内容`addTextBox`相同。采用字符串参数设置的文本`TextRange`。

下面的代码示例演示如何创建带有文本 "Hello!" 的文本框。

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

可以将形状组合在一起。 这样一来, 用户可以将它们视为定位、调整大小和其他相关任务的单个实体。 [ShapeGroup](/javascript/api/excel/excel.shapegroup)是一种类型的`Shape`, 因此加载项将该组视为单个形状。

下面的代码示例演示组合在一起的三个形状。 后续代码示例显示了形状组被移至右侧50像素。

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
> 组中的单个形状是通过`ShapeGroup.shapes`属性引用的, 该属性的类型为[GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection)。 分组后, 将无法再通过工作表的形状集合访问它们。 例如, 如果您的工作表中有三个形状, 并且它们都组合在一起, `shapes.getCount`则工作表的方法将返回一个计数为1。

## <a name="export-shapes-as-images"></a>将形状导出为图像

任何`Shape`对象都可以转换为图像。 [getAsImage](/javascript/api/excel/excel.shape#getasimage-format-)返回 base64 编码的字符串。 图像的格式被指定为传递给[](/javascript/api/excel/excel.pictureformat) `getAsImage`的 PictureFormat 枚举。

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

使用`Shape`对象的`delete`方法从工作表中删除形状。 无需任何其他元数据。

下面的代码示例从**MyWorksheet**中删除所有形状。

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
