---
title: 使用 JavaScript API Excel形状
description: 了解如何Excel形状定义为位于绘图层上的任何Excel。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: eeb6a1f76c839e4b550662b28b717bfd1bcca4e8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349446"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a><span data-ttu-id="c0932-103">使用 JavaScript API Excel形状</span><span class="sxs-lookup"><span data-stu-id="c0932-103">Work with shapes using the Excel JavaScript API</span></span>

<span data-ttu-id="c0932-104">Excel形状定义为位于绘图层上的任何Excel。</span><span class="sxs-lookup"><span data-stu-id="c0932-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="c0932-105">这意味着单元格之外任何内容都是形状。</span><span class="sxs-lookup"><span data-stu-id="c0932-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="c0932-106">本文介绍如何将几何形状、线条和图像与 [Shape](/javascript/api/excel/excel.shape) 和 [ShapeCollection](/javascript/api/excel/excel.shapecollection) API 结合使用。</span><span class="sxs-lookup"><span data-stu-id="c0932-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="c0932-107">[图](/javascript/api/excel/excel.chart)在其自己的文章使用[JavaScript API](excel-add-ins-charts.md)处理图表Excel介绍。</span><span class="sxs-lookup"><span data-stu-id="c0932-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

<span data-ttu-id="c0932-108">下图显示了形成温度计的形状。</span><span class="sxs-lookup"><span data-stu-id="c0932-108">The following image shows shapes which form a thermometer.</span></span>
<span data-ttu-id="c0932-109">![作为温度计的形状Excel的图像。](../images/excel-shapes.png)</span><span class="sxs-lookup"><span data-stu-id="c0932-109">![Image of a thermometer made as an Excel shape.](../images/excel-shapes.png)</span></span>

## <a name="create-shapes"></a><span data-ttu-id="c0932-110">创建形状</span><span class="sxs-lookup"><span data-stu-id="c0932-110">Create shapes</span></span>

<span data-ttu-id="c0932-111">形状通过工作表的形状集合创建并存储在 `Worksheet.shapes` () 。</span><span class="sxs-lookup"><span data-stu-id="c0932-111">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="c0932-112">`ShapeCollection` 为此 `.add*` ，有几个方法。</span><span class="sxs-lookup"><span data-stu-id="c0932-112">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="c0932-113">所有形状在添加到集合时都有为它们生成的名称和 ID。</span><span class="sxs-lookup"><span data-stu-id="c0932-113">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="c0932-114">它们分别为 `name` 和 `id` 属性。</span><span class="sxs-lookup"><span data-stu-id="c0932-114">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="c0932-115">`name` 可通过外接程序设置 ，以使用 方法轻松 `ShapeCollection.getItem(name)` 检索。</span><span class="sxs-lookup"><span data-stu-id="c0932-115">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="c0932-116">以下类型的形状是使用关联方法添加的。</span><span class="sxs-lookup"><span data-stu-id="c0932-116">The following types of shapes are added using the associated method.</span></span>

| <span data-ttu-id="c0932-117">形状</span><span class="sxs-lookup"><span data-stu-id="c0932-117">Shape</span></span> | <span data-ttu-id="c0932-118">Add 方法</span><span class="sxs-lookup"><span data-stu-id="c0932-118">Add Method</span></span> | <span data-ttu-id="c0932-119">签名</span><span class="sxs-lookup"><span data-stu-id="c0932-119">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="c0932-120">几何形状</span><span class="sxs-lookup"><span data-stu-id="c0932-120">Geometric Shape</span></span> | [<span data-ttu-id="c0932-121">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="c0932-121">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="c0932-122">图像 (JPEG 或 PNG) </span><span class="sxs-lookup"><span data-stu-id="c0932-122">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="c0932-123">addImage</span><span class="sxs-lookup"><span data-stu-id="c0932-123">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="c0932-124">折线图</span><span class="sxs-lookup"><span data-stu-id="c0932-124">Line</span></span> | [<span data-ttu-id="c0932-125">addLine</span><span class="sxs-lookup"><span data-stu-id="c0932-125">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="c0932-126">SVG</span><span class="sxs-lookup"><span data-stu-id="c0932-126">SVG</span></span> | [<span data-ttu-id="c0932-127">addSvg</span><span class="sxs-lookup"><span data-stu-id="c0932-127">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="c0932-128">文本框</span><span class="sxs-lookup"><span data-stu-id="c0932-128">Text Box</span></span> | [<span data-ttu-id="c0932-129">addTextBox</span><span class="sxs-lookup"><span data-stu-id="c0932-129">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="c0932-130">几何形状</span><span class="sxs-lookup"><span data-stu-id="c0932-130">Geometric shapes</span></span>

<span data-ttu-id="c0932-131">使用 创建几何形状 `ShapeCollection.addGeometricShape` 。</span><span class="sxs-lookup"><span data-stu-id="c0932-131">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="c0932-132">该方法采用 [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) 枚举作为参数。</span><span class="sxs-lookup"><span data-stu-id="c0932-132">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="c0932-133">下面的代码示例创建一个名为 **"Square"** 的 150x150 像素矩形，该矩形从工作表的顶部和左侧放置 100 像素。</span><span class="sxs-lookup"><span data-stu-id="c0932-133">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="c0932-134">图像</span><span class="sxs-lookup"><span data-stu-id="c0932-134">Images</span></span>

<span data-ttu-id="c0932-135">JPEG、PNG 和 SVG 图像可以作为形状插入到工作表中。</span><span class="sxs-lookup"><span data-stu-id="c0932-135">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="c0932-136">`ShapeCollection.addImage`方法将 base64 编码的字符串作为参数。</span><span class="sxs-lookup"><span data-stu-id="c0932-136">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="c0932-137">这是字符串形式的 JPEG 或 PNG 图像。</span><span class="sxs-lookup"><span data-stu-id="c0932-137">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="c0932-138">`ShapeCollection.addSvg` 即使此参数是定义图形的 XML，也采用字符串。</span><span class="sxs-lookup"><span data-stu-id="c0932-138">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="c0932-139">下面的代码示例演示 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 作为字符串加载的图像文件。</span><span class="sxs-lookup"><span data-stu-id="c0932-139">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="c0932-140">该字符串具有元数据"base64"，该元数据在形状创建之前被删除。</span><span class="sxs-lookup"><span data-stu-id="c0932-140">The string has the metadata "base64," removed before the shape is created.</span></span>

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

### <a name="lines"></a><span data-ttu-id="c0932-141">Lines</span><span class="sxs-lookup"><span data-stu-id="c0932-141">Lines</span></span>

<span data-ttu-id="c0932-142">使用 创建行 `ShapeCollection.addLine` 。</span><span class="sxs-lookup"><span data-stu-id="c0932-142">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="c0932-143">该方法需要线条起点和终点的左边距和上边距。</span><span class="sxs-lookup"><span data-stu-id="c0932-143">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="c0932-144">它还需要 [一个 ConnectorType](/javascript/api/excel/excel.connectortype) 枚举来指定终结点之间的直线连接。</span><span class="sxs-lookup"><span data-stu-id="c0932-144">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="c0932-145">下面的代码示例在工作表上创建一条直线。</span><span class="sxs-lookup"><span data-stu-id="c0932-145">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c0932-146">线条可以连接到其他 Shape 对象。</span><span class="sxs-lookup"><span data-stu-id="c0932-146">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="c0932-147">和 `connectBeginShape` `connectEndShape` 方法将线条的开头和结尾附加到指定连接点处的形状。</span><span class="sxs-lookup"><span data-stu-id="c0932-147">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="c0932-148">这些点的位置因形状而异，但 可用于确保您的外接程序不会连接到外部 `Shape.connectionSiteCount` 的点。</span><span class="sxs-lookup"><span data-stu-id="c0932-148">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="c0932-149">线条使用 和 方法断开与任何附加 `disconnectBeginShape` `disconnectEndShape` 形状的连接。</span><span class="sxs-lookup"><span data-stu-id="c0932-149">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="c0932-150">下面的代码示例将 **"MyLine"** 行连接到名为 **"LeftShape"** 和 **"RightShape"的两个形状**。</span><span class="sxs-lookup"><span data-stu-id="c0932-150">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="c0932-151">移动形状并调整形状大小</span><span class="sxs-lookup"><span data-stu-id="c0932-151">Move and resize shapes</span></span>

<span data-ttu-id="c0932-152">形状位于工作表的顶部。</span><span class="sxs-lookup"><span data-stu-id="c0932-152">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="c0932-153">它们的位置由 和 `left` `top` 属性定义。</span><span class="sxs-lookup"><span data-stu-id="c0932-153">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="c0932-154">它们充当工作表各自边缘的边距，[0， 0] 为左上角。</span><span class="sxs-lookup"><span data-stu-id="c0932-154">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="c0932-155">可以使用 和 方法直接设置它们，也可以从当前位置 `incrementLeft` `incrementTop` 进行调整。</span><span class="sxs-lookup"><span data-stu-id="c0932-155">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="c0932-156">此外，还通过此方式建立形状相对于默认位置的旋转量，该属性为绝对量，而方法用于调整现有 `rotation` `incrementRotation` 旋转量。</span><span class="sxs-lookup"><span data-stu-id="c0932-156">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="c0932-157">形状相对于其他形状的深度由 属性 `zorderPosition` 定义。</span><span class="sxs-lookup"><span data-stu-id="c0932-157">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="c0932-158">这是使用 方法 `setZOrder` 设置的，该方法采用 [ShapeZOrder](/javascript/api/excel/excel.shapezorder)。</span><span class="sxs-lookup"><span data-stu-id="c0932-158">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="c0932-159">`setZOrder` 调整当前形状相对于其他形状的排序。</span><span class="sxs-lookup"><span data-stu-id="c0932-159">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="c0932-160">外接程序具有多个选项，用于更改形状的高度和宽度。</span><span class="sxs-lookup"><span data-stu-id="c0932-160">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="c0932-161">如果设置 `height` 或 `width` 属性，则更改指定维度而不更改其他维度。</span><span class="sxs-lookup"><span data-stu-id="c0932-161">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="c0932-162">和 `scaleHeight` `scaleWidth` 根据所提供的 [ShapeScaleType](/javascript/api/excel/excel.shapescaletype) 属性的值调整形状相对于当前大小或原始 (调整形状) 。</span><span class="sxs-lookup"><span data-stu-id="c0932-162">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="c0932-163">可选的 [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) 参数指定形状缩放位置 (左上角、中间或右下角缩放) 。</span><span class="sxs-lookup"><span data-stu-id="c0932-163">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="c0932-164">如果 `lockAspectRatio` 该属性为 **true，** 则缩放方法通过同时调整其他尺寸来保持形状的当前纵横比。</span><span class="sxs-lookup"><span data-stu-id="c0932-164">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="c0932-165">对 和 `height` `width` 属性的直接更改仅影响该属性，而 `lockAspectRatio` 不管属性值如何。</span><span class="sxs-lookup"><span data-stu-id="c0932-165">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="c0932-166">下面的代码示例显示一个形状，其缩放比例为其原始大小的 1.25 倍，并旋转 30 度。</span><span class="sxs-lookup"><span data-stu-id="c0932-166">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="c0932-167">形状中的文本</span><span class="sxs-lookup"><span data-stu-id="c0932-167">Text in shapes</span></span>

<span data-ttu-id="c0932-168">几何形状可以包含文本。</span><span class="sxs-lookup"><span data-stu-id="c0932-168">Geometric Shapes can contain text.</span></span> <span data-ttu-id="c0932-169">形状具有 `textFrame` [TextFrame 类型的属性](/javascript/api/excel/excel.textframe)。</span><span class="sxs-lookup"><span data-stu-id="c0932-169">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="c0932-170">对象 `TextFrame` 管理文本显示选项 (如边距和文本溢出) 。</span><span class="sxs-lookup"><span data-stu-id="c0932-170">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="c0932-171">`TextFrame.textRange` 是文本内容和字体设置的 [TextRange](/javascript/api/excel/excel.textrange) 对象。</span><span class="sxs-lookup"><span data-stu-id="c0932-171">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="c0932-172">下面的代码示例创建一个名为"Wave"的几何形状，其文本为"Shape Text"。</span><span class="sxs-lookup"><span data-stu-id="c0932-172">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="c0932-173">它还调整形状和文本颜色，以及将文本的水平对齐方式设置到中心。</span><span class="sxs-lookup"><span data-stu-id="c0932-173">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="c0932-174">`addTextBox`方法创建 `ShapeCollection` 具有 `GeometricShape` 白色 `Rectangle` 背景和黑色文本的 类型。</span><span class="sxs-lookup"><span data-stu-id="c0932-174">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="c0932-175">这与"插入"选项卡 **Excel"** 文本框"按钮 **所创建** 的内容相同。 `addTextBox`采用字符串参数来设置 的文本 `TextRange` 。</span><span class="sxs-lookup"><span data-stu-id="c0932-175">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="c0932-176">下面的代码示例演示了如何创建文本为"Hello！"的文本框。</span><span class="sxs-lookup"><span data-stu-id="c0932-176">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="c0932-177">形状组</span><span class="sxs-lookup"><span data-stu-id="c0932-177">Shape groups</span></span>

<span data-ttu-id="c0932-178">形状可以组合在一起。</span><span class="sxs-lookup"><span data-stu-id="c0932-178">Shapes can be grouped together.</span></span> <span data-ttu-id="c0932-179">这允许用户将它们视为用于定位、调整大小和其他相关任务的单个实体。</span><span class="sxs-lookup"><span data-stu-id="c0932-179">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="c0932-180">[ShapeGroup](/javascript/api/excel/excel.shapegroup)是 的一种类型，因此加载项将组 `Shape` 视为单个形状。</span><span class="sxs-lookup"><span data-stu-id="c0932-180">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="c0932-181">下面的代码示例显示了组合在一起的三个形状。</span><span class="sxs-lookup"><span data-stu-id="c0932-181">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="c0932-182">后续代码示例显示形状组向右移动 50 像素。</span><span class="sxs-lookup"><span data-stu-id="c0932-182">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="c0932-183">组内的单个形状通过属性引用，该属性的类型 `ShapeGroup.shapes` 为 [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection)。</span><span class="sxs-lookup"><span data-stu-id="c0932-183">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="c0932-184">分组后，无法再通过工作表的形状集合访问它们。</span><span class="sxs-lookup"><span data-stu-id="c0932-184">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="c0932-185">例如，如果工作表有三个形状，并且它们全部组合在一起，则工作表 `shapes.getCount` 的方法将返回计数 1。</span><span class="sxs-lookup"><span data-stu-id="c0932-185">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="c0932-186">将形状导出为图像</span><span class="sxs-lookup"><span data-stu-id="c0932-186">Export shapes as images</span></span>

<span data-ttu-id="c0932-187">任何 `Shape` 对象都可以转换为图像。</span><span class="sxs-lookup"><span data-stu-id="c0932-187">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="c0932-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) 返回 base64 编码的字符串。</span><span class="sxs-lookup"><span data-stu-id="c0932-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="c0932-189">图像的格式指定为传递给 的 [PictureFormat](/javascript/api/excel/excel.pictureformat) 枚举 `getAsImage` 。</span><span class="sxs-lookup"><span data-stu-id="c0932-189">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="c0932-190">删除形状</span><span class="sxs-lookup"><span data-stu-id="c0932-190">Delete shapes</span></span>

<span data-ttu-id="c0932-191">使用对象的方法从工作表 `Shape` 中删除 `delete` 形状。</span><span class="sxs-lookup"><span data-stu-id="c0932-191">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="c0932-192">无需其他元数据。</span><span class="sxs-lookup"><span data-stu-id="c0932-192">No other metadata is needed.</span></span>

<span data-ttu-id="c0932-193">下面的代码示例从 **MyWorksheet 中删除所有形状**。</span><span class="sxs-lookup"><span data-stu-id="c0932-193">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c0932-194">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c0932-194">See also</span></span>

- [<span data-ttu-id="c0932-195">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="c0932-195">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="c0932-196">使用 Excel JavaScript API 处理图表</span><span class="sxs-lookup"><span data-stu-id="c0932-196">Work with charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
