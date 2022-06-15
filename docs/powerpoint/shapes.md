---
title: 使用 PowerPoint JavaScript API 处理形状
description: 了解如何在PowerPoint幻灯片上添加、删除和设置形状的格式。
ms.date: 06/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f314cfebb26450e79dbabe1e65ac9e4c8fe9799
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091102"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api"></a>使用 PowerPoint JavaScript API 处理形状

本文介绍如何将几何形状、线条和文本框与 [Shape](/javascript/api/powerpoint/powerpoint.shape) 和 [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) API 结合使用。

## <a name="create-shapes"></a>创建形状

形状是通过幻灯片的形状集合 () `slide.shapes` 创建和存储的。 `ShapeCollection` 有几 `.add*` 种方法用于此目的。 所有形状在添加到集合时都会为它们生成名称和 ID。 这些是分别的 `name` 属性和 `id` 属性。 `name` 可由加载项设置。

### <a name="geometric-shapes"></a>几何形状

使用其中一个重载 `ShapeCollection.addGeometricShape`创建几何形状。 第一个参数是 [GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) 枚举或等效于枚举值之一的字符串。 有一个可选的第二个参数，类型 [为 ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) ，可以指定形状的初始大小及其相对于幻灯片顶部和左侧的位置（以磅为单位）。 也可以在创建形状后设置这些属性。

下面的代码示例创建一个名为 **“Square”** 的矩形，该矩形从幻灯片的顶部和左侧放置 100 磅。 该方法返回一个 `Shape` 对象。

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

使用其中一个重载 `ShapeCollection.addLine`创建行。 第一个参数是 [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) 枚举或等效于枚举值之一的字符串，用于指定行在终结点之间的串串方式。 有一个可选的第二个参数，类型 [为 ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) ，可以指定行的起始点和终点。 也可以在创建形状后设置这些属性。 该方法返回一个 `Shape` 对象。

> [!NOTE]
> 当形状为线条时，`top``left`和`ShapeAddOptions`对象的`Shape`属性指定相对于幻灯片的上边缘和左边缘的线条的起始点。 和`height``width`属性指定 *相对于起点* 的行的终结点。 因此，相对于幻灯片顶部和左边缘的终结点 (`top` + `height`)  () 。`left` + `width` 所有属性的度量单位为磅，允许负值。

以下代码示例在幻灯片上创建直线。

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

使用 [addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) 方法创建文本框。 第一个参数是最初应显示在框中的文本。 有一个可选的第二个参数类型 [为 ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) ，可以指定文本框的初始大小及其相对于幻灯片顶部和左侧的位置。 也可以在创建形状后设置这些属性。

以下代码示例演示如何在第一张幻灯片上创建文本框。

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

## <a name="move-and-resize-shapes"></a>移动和调整形状的大小

形状位于幻灯片顶部。 它们的位置由 `left` 属性和 `top` 属性定义。 这些作为幻灯片各自边缘的边距，以磅为单位，`left: 0``top: 0`左上角。 形状大小由`height``width`和属性指定。 代码可以通过重置这些属性来移动或调整形状的大小。  (当形状为线条时，这些属性的含义略有不同。 请参阅 [Lines](#lines).) 

## <a name="text-in-shapes"></a>形状中的文本

几何形状可以包含文本。 形状的 `textFrame` 属性类型为 [TextFrame](/javascript/api/powerpoint/powerpoint.textframe)。 该 `TextFrame` 对象管理文本显示选项 (如边距和文本溢出) 。 `TextFrame.textRange` 是包含文本内容和字体设置的 [TextRange](/javascript/api/powerpoint/powerpoint.textrange) 对象。

下面的代码示例创建一个名为 **“大括号”** 的几何形状，其文本 **为“形状文本”。** 它还调整形状和文本颜色，并设置文本与中心的垂直对齐方式。

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

使用对象`delete`的方法从幻灯片中`Shape`删除形状。

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
