---
title: 使用 Excel JavaScript API 处理形状
description: 了解 Excel 如何将形状定义为位于 Excel 绘图层上的任何对象。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 507ae05b570e7eef4f3bf5560ca47c1bfbd40f9f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889595"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理形状

Excel 将形状定义为位于 Excel 绘图层上的任何对象。 这意味着单元格之外的任何内容都是形状。 本文介绍如何将几何形状、线条和图像与 [Shape](/javascript/api/excel/excel.shape) 和 [ShapeCollection](/javascript/api/excel/excel.shapecollection) API 结合使用。 [图表](/javascript/api/excel/excel.chart) 在其自己的文章“ [使用 Excel JavaScript API 处理图表](excel-add-ins-charts.md)”中进行了介绍。

下图显示了构成温度计的形状。
![作为 Excel 形状制作的温度计图像。](../images/excel-shapes.png)

## <a name="create-shapes"></a>创建形状

形状是通过工作表的形状集合 () `Worksheet.shapes` 创建和存储的。 `ShapeCollection` 有几 `.add*` 种方法用于此目的。 所有形状在添加到集合时都会为它们生成名称和 ID。 这些是分别的 `name` 属性和 `id` 属性。 `name` 可由外接程序设置，以便使用该 `ShapeCollection.getItem(name)` 方法轻松检索。

使用关联的方法添加以下类型的形状。

| 形状 | Add 方法 | 签名 |
|-------|------------|-----------|
| 几何形状 | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 图像 (JPEG 或 PNG)  | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| 折线图 | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| 文本框 | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>几何形状

使用 a0 创建 `ShapeCollection.addGeometricShape`几何形状。 该方法采用 [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) 枚举作为参数。

下面的代码示例创建一个名为 **“Square”** 的 150x150 像素矩形，该矩形从工作表的顶部和左侧放置 100 像素。

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### <a name="images"></a>图像

JPEG、PNG 和 SVG 映像可以作为形状插入到工作表中。 该 `ShapeCollection.addImage` 方法采用 base64 编码的字符串作为参数。 这是字符串形式的 JPEG 或 PNG 映像。 `ShapeCollection.addSvg` 也采用字符串，尽管此参数是定义图形的 XML。

下面的代码示例显示 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 作为字符串加载的图像文件。 字符串在创建形状之前删除了元数据“base64”。

```js
// This sample creates an image as a Shape object in the worksheet.
let myFile = document.getElementById("selectedFile");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        let startIndex = reader.result.toString().indexOf("base64,");
        let myBase64 = reader.result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getItem("MyWorksheet");
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

使用 a0 创建 `ShapeCollection.addLine`行。 该方法需要线的起始点和终点的左侧和顶部边距。 还需要一个 [ConnectorType](/javascript/api/excel/excel.connectortype) 枚举来指定终结点之间的行串联方式。 下面的代码示例在工作表上创建一条直线。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

线条可以连接到其他 Shape 对象。 和`connectEndShape`方法将`connectBeginShape`线条的开头和结尾附加到指定连接点处的形状。 这些点的位置因形状而异，但 `Shape.connectionSiteCount` 可用于确保加载项不会连接到超出边界的点。 使用和方法，线条与`disconnectBeginShape``disconnectEndShape`任何附加的形状断开连接。

以下代码示例将 **“MyLine”** 行连接到两个名为 **“LeftShape”** 和 **“RightShape”的** 形状。

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>移动和调整形状的大小

形状位于工作表顶部。 它们的位置由属性 `left` 和 `top` 属性定义。 这些充当工作表各自边缘的边距，其中 [0， 0] 是左上角。 这些可以直接设置或调整从其当前位置与 `incrementLeft` 和 `incrementTop` 方法。 形状从默认位置旋转多少也是以这种方式建立的 `rotation` ，属性是绝对量， `incrementRotation` 方法调整现有旋转。

形状相对于其他形状的深度由 `zorderPosition` 属性定义。 这是使用 `setZOrder` 采用 [ShapeZOrder](/javascript/api/excel/excel.shapezorder) 的方法设置的。 `setZOrder` 调整当前形状相对于其他形状的顺序。

外接程序有几个选项用于更改形状的高度和宽度。 `height`设置或`width`属性会更改指定的维度，而不会更改其他维度。 `scaleWidth`根据`scaleHeight`提供的 [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)) 的值，根据当前或原始大小 (调整形状各自的尺寸。 可选 [的 ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) 参数指定形状从中缩放 (左上角、中间或右下角) 。 如果属性 `lockAspectRatio` 是 `true`，则缩放方法还通过调整其他维度来保持形状的当前纵横比。

> [!NOTE]
> 对属性和属性的直接`height`更改仅影响该属性，而不考虑`lockAspectRatio`属性的`width`值。

下面的代码示例显示一个形状被缩放到其原始大小的 1.25 倍，并旋转了 30 度。

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");

    let shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);

    await context.sync();
});
```

## <a name="text-in-shapes"></a>形状中的文本

几何形状可以包含文本。 形状的 `textFrame` 属性类型为 [TextFrame](/javascript/api/excel/excel.textframe)。 该 `TextFrame` 对象管理文本显示选项 (如边距和文本溢出) 。 `TextFrame.textRange` 是包含文本内容和字体设置的 [TextRange](/javascript/api/excel/excel.textrange) 对象。

下面的代码示例创建一个名为“Wave”的几何形状，其文本为“形状文本”。 它还调整形状和文本颜色，并将文本的水平对齐方式设置为中心。

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;

    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");

    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;

    await context.sync();
});
```

创建`addTextBox``GeometricShape`具有白色背景和黑色文本的类型的`Rectangle`方法`ShapeCollection`。 这与 Excel 的“**插入**”选项卡上的 **“文本框**”按钮创建的内容相同。 `addTextBox`使用字符串参数设置文本`TextRange`。

下面的代码示例演示如何创建文本框，文本为“Hello！”。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="shape-groups"></a>形状组

形状可以组合在一起。 这允许用户将其视为用于定位、调整大小和其他相关任务的单个实体。 [ShapeGroup](/javascript/api/excel/excel.shapegroup) 是一种类型`Shape`，因此加载项将组视为单个形状。

下面的代码示例显示将三个形状分组在一起。 后续代码示例显示要移动到右侧 50 像素的形状组。

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let square = shapes.getItem("Square");
    let pentagon = shapes.getItem("Pentagon");
    let octagon = shapes.getItem("Octagon");

    let shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    await context.sync();
});

// This sample moves the previously created shape group to the right by 50 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    await context.sync();
});
```

> [!IMPORTANT]
> 组中的单个形状通过 `ShapeGroup.shapes` 类型为 [GroupShapeCollection 的](/javascript/api/excel/excel.groupshapecollection)属性进行引用。 分组后，不再可通过工作表的形状集合访问它们。 例如，如果工作表有三个形状，并且它们都组合在一起，则工作表 `shapes.getCount` 的方法将返回 1 的计数。

## <a name="export-shapes-as-images"></a>将形状导出为图像

任何 `Shape` 对象都可以转换为图像。 [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) 返回 base64 编码的字符串。 图像的格式指定为传递给`getAsImage`的 [PictureFormat](/javascript/api/excel/excel.pictureformat) 枚举。

```js
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shape = shapes.getItem("Image");
    let stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
});
```

## <a name="delete-shapes"></a>删除形状

使用对象`delete`的方法从工作表中`Shape`删除形状。 不需要其他元数据。

以下代码示例从 **MyWorksheet** 中删除所有形状。

```js
// This deletes all the shapes from "MyWorksheet".
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");
    let shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    
    await context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)
- [使用 Excel JavaScript API 处理图表](excel-add-ins-charts.md)
