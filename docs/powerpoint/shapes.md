---
title: 使用 JavaScript API PowerPoint形状
description: 了解如何在幻灯片上添加、删除形状和PowerPoint格式。
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2c7eb7a1770f807878320369951faa7d0ddc873c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340482"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api-preview"></a>使用 JavaScript API PowerPoint预览 (形状) 

本文介绍如何将几何形状、线条和文本框与 Shape 和 [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) API 结合使用[](/javascript/api/powerpoint/powerpoint.shape)。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="create-shapes"></a>创建形状

形状通过幻灯片的形状集合创建并存储在幻灯片中 () `slide.shapes` 。 `ShapeCollection` 为此， `.add*` 有几个方法。 所有形状在添加到集合时都有为它们生成的名称和 ID。 它们分别为 `name` 和 `id` 属性。 `name` 可通过加载项进行设置。

### <a name="geometric-shapes"></a>几何形状

使用 的重载之一创建几何形状 `ShapeCollection.addGeometricShape`。 第一个参数是 [GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) 枚举或等价于该枚举值之一的字符串。 有一个 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 类型的可选第二个参数，该参数可指定形状的初始大小及其相对于幻灯片顶部和左侧的位置（以点为单位）。 也可以创建形状后设置这些属性。

下面的代码示例创建一个名为" **Square"** 的矩形，该矩形位于从幻灯片的上边缘和左边 100 个点处。 方法返回对象 `Shape` 。

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    await context.sync();
});
```

### <a name="lines"></a>Lines

使用 的重载之一创建行 `ShapeCollection.addLine`。 第一个参数是 [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) 枚举或等价于枚举值之一的字符串，用于指定线在终结点之间如何相互连接。 有一个 [类型为 ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 的可选第二个参数，该参数可指定线条的起始点和终点。 也可以创建形状后设置这些属性。 方法返回对象 `Shape` 。

> [!NOTE]
> 当形状是线条时 `top` `left` `Shape` `ShapeAddOptions` ，和 对象的 和 属性指定线条相对于幻灯片的上边缘和左边缘的起始点。 和 `height` `width` 属性指定线条相对于 *起点的终点*。 因此，相对于幻灯片的上边缘和左 `top` + `height` 边缘的终点 ()  (`left` + `width`) 。 所有属性的度量单位都是点，允许使用负值。

下面的代码示例在幻灯片上创建一条直线。

```js
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    await context.sync();
});
```

### <a name="text-boxes"></a>文本框

使用 [addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) 方法创建一个文本框。 第一个参数是最初应显示在框中的文本。 有一个 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 类型的可选第二个参数，该参数可以指定文本框的初始大小及其相对于幻灯片顶部和左侧的位置。 也可以创建形状后设置这些属性。

下面的代码示例演示如何创建第一张幻灯片上的文本框。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>移动形状并调整形状大小

形状位于幻灯片顶部。 它们的位置由 和 `left` 属性 `top` 定义。 它们充当幻灯片`left: 0``top: 0`各自边缘的边距（以点为单位）以及左上角。 形状大小由 和 属性`height``width`指定。 您的代码可以通过重置这些属性来移动或调整形状的大小。  (形状是线条时，这些属性的含义略有不同。 请参阅 [Lines](#lines).) 

## <a name="text-in-shapes"></a>形状中的文本

几何形状可以包含文本。 形状具有 TextFrame `textFrame` [类型的属性](/javascript/api/powerpoint/powerpoint.textframe)。 对象 `TextFrame` 管理文本显示选项 (如边距和文本溢出) 。 `TextFrame.textRange` 是文本内容和字体设置的 [TextRange](/javascript/api/powerpoint/powerpoint.textrange) 对象。

下面的代码示例创建一个名为 **"大** 括号"的几何形状，其文本为 **"Shape text"**。 它还调整形状和文本颜色，以及将文本的垂直对齐方式设置到中心。

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    await context.sync();
});
```

## <a name="delete-shapes"></a>删除形状

使用对象的方法从幻灯片 `Shape` 中删除形状 `delete` 。

下面的代码示例演示如何删除形状。

```js
await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const sheet = context.presentation.slides.getItemAt(0);
    const shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
        
    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    await context.sync();
});
```
