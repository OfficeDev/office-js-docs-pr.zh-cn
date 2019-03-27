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
# <a name="work-with-shapes-using-the-excel-javascript-api-preview"></a><span data-ttu-id="494ce-102">使用 Excel JavaScript API 处理形状 (预览)</span><span class="sxs-lookup"><span data-stu-id="494ce-102">Work with Shapes using the Excel JavaScript API (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="494ce-103">本文中讨论的 api 目前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="494ce-103">The APIs discussed in this article are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="494ce-104">excel 将形状定义为位于 Excel 绘图层的任何对象。</span><span class="sxs-lookup"><span data-stu-id="494ce-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="494ce-105">这意味着单元格之外的任何内容都是形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="494ce-106">本文介绍如何将几何形状、线条和图像与 [Shape]/javascript/api/excel/excel.shape) 和[ShapeCollection](/javascript/api/excel/excel.shapecollection) api 结合使用。</span><span class="sxs-lookup"><span data-stu-id="494ce-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape]/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="494ce-107">[图表](/javascript/api/excel/excel.chart)在其自己的文章中进行了介绍, [使用 Excel JavaScript API 处理图表]] (Excel 外接程序-charts.md))。</span><span class="sxs-lookup"><span data-stu-id="494ce-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with Charts using the Excel JavaScript API]](excel-add-ins-charts.md)).</span></span>

## <a name="create-shapes"></a><span data-ttu-id="494ce-108">创建形状</span><span class="sxs-lookup"><span data-stu-id="494ce-108">Create shapes</span></span>

<span data-ttu-id="494ce-109">形状是通过工作表的形状集合 (`Worksheet.shapes`) 创建和存储的。</span><span class="sxs-lookup"><span data-stu-id="494ce-109">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="494ce-110">`ShapeCollection`有几`.add*`种方法可以实现此目的。</span><span class="sxs-lookup"><span data-stu-id="494ce-110">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="494ce-111">在将所有形状添加到集合中时, 都会为它们生成名称和 id。</span><span class="sxs-lookup"><span data-stu-id="494ce-111">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="494ce-112">它们分别是`name`和`id`属性。</span><span class="sxs-lookup"><span data-stu-id="494ce-112">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="494ce-113">`name`可通过外接程序进行设置以便使用`ShapeCollection.getItem(name)`方法轻松检索。</span><span class="sxs-lookup"><span data-stu-id="494ce-113">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="494ce-114">使用关联的方法添加以下类型的形状:</span><span class="sxs-lookup"><span data-stu-id="494ce-114">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="494ce-115">Shape</span><span class="sxs-lookup"><span data-stu-id="494ce-115">Shape</span></span> | <span data-ttu-id="494ce-116">Add 方法</span><span class="sxs-lookup"><span data-stu-id="494ce-116">Add Method</span></span> | <span data-ttu-id="494ce-117">签名</span><span class="sxs-lookup"><span data-stu-id="494ce-117">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="494ce-118">几何形状</span><span class="sxs-lookup"><span data-stu-id="494ce-118">Geometric Shape</span></span> | [<span data-ttu-id="494ce-119">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="494ce-119">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="494ce-120">图像 (JPEG 或 PNG)</span><span class="sxs-lookup"><span data-stu-id="494ce-120">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="494ce-121">addImage</span><span class="sxs-lookup"><span data-stu-id="494ce-121">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="494ce-122">Line</span><span class="sxs-lookup"><span data-stu-id="494ce-122">Line</span></span> | [<span data-ttu-id="494ce-123">addLine</span><span class="sxs-lookup"><span data-stu-id="494ce-123">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="494ce-124">SVG</span><span class="sxs-lookup"><span data-stu-id="494ce-124">SVG</span></span> | [<span data-ttu-id="494ce-125">addSvg</span><span class="sxs-lookup"><span data-stu-id="494ce-125">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="494ce-126">文本框</span><span class="sxs-lookup"><span data-stu-id="494ce-126">Text Box</span></span> | [<span data-ttu-id="494ce-127">addTextBox</span><span class="sxs-lookup"><span data-stu-id="494ce-127">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="494ce-128">几何形状</span><span class="sxs-lookup"><span data-stu-id="494ce-128">Geometric shapes</span></span>

<span data-ttu-id="494ce-129">将使用`ShapeCollection.addGeometricShape`创建一个几何形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-129">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="494ce-130">该方法采用[GeometricShapeType](/javascript/api/excel/excel.geometricshapetype)枚举作为参数。</span><span class="sxs-lookup"><span data-stu-id="494ce-130">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="494ce-131">下面的代码示例创建一个名为 **"** 150x150" 的像素矩形, 该矩形在工作表的顶部和左侧位置为100像素。</span><span class="sxs-lookup"><span data-stu-id="494ce-131">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="494ce-132">图像</span><span class="sxs-lookup"><span data-stu-id="494ce-132">Images</span></span>

<span data-ttu-id="494ce-133">JPEG、PNG 和 SVG 图像可以作为形状插入到工作表中。</span><span class="sxs-lookup"><span data-stu-id="494ce-133">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="494ce-134">该`ShapeCollection.addImage`方法采用 base64 编码的字符串作为参数。</span><span class="sxs-lookup"><span data-stu-id="494ce-134">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="494ce-135">这是字符串形式的 JPEG 或 PNG 图像。</span><span class="sxs-lookup"><span data-stu-id="494ce-135">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="494ce-136">`ShapeCollection.addSvg`也采用字符串, 但此参数是用于定义图形的 XML。</span><span class="sxs-lookup"><span data-stu-id="494ce-136">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="494ce-137">下面的代码示例显示了[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader)作为字符串加载的图像文件。</span><span class="sxs-lookup"><span data-stu-id="494ce-137">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="494ce-138">字符串中包含元数据 "base64", 在创建形状之前将其删除。</span><span class="sxs-lookup"><span data-stu-id="494ce-138">The string has the metadata "base64," removed before the shape is created.</span></span>

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

### <a name="lines"></a><span data-ttu-id="494ce-139">Lines</span><span class="sxs-lookup"><span data-stu-id="494ce-139">Lines</span></span>

<span data-ttu-id="494ce-140">将使用`ShapeCollection.addLine`创建的行。</span><span class="sxs-lookup"><span data-stu-id="494ce-140">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="494ce-141">该方法需要线条的起始点和结束点的左边距和上边距。</span><span class="sxs-lookup"><span data-stu-id="494ce-141">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="494ce-142">它还采用[ConnectorType](/javascript/api/excel/excel.connectortype)枚举来指定行在终结点之间的 contorts 方式。</span><span class="sxs-lookup"><span data-stu-id="494ce-142">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="494ce-143">下面的代码示例在工作表上创建一条直线。</span><span class="sxs-lookup"><span data-stu-id="494ce-143">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="494ce-144">可以将线条连接到其他 Shape 对象。</span><span class="sxs-lookup"><span data-stu-id="494ce-144">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="494ce-145">`connectBeginShape`和`connectEndShape`方法将行的开头和结尾附加到指定连接点处的形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-145">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="494ce-146">这些点的位置因形状而异, 但`Shape.connectionSiteCount`可用于确保外接程序不会连接到超出边界的点。</span><span class="sxs-lookup"><span data-stu-id="494ce-146">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="494ce-147">使用`disconnectBeginShape`和`disconnectEndShape`方法将线条与任何附加的形状断开连接。</span><span class="sxs-lookup"><span data-stu-id="494ce-147">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="494ce-148">下面的代码示例将 **"MyLine"** 行连接到名为 **"LeftShape"** 和 **"RightShape"** 的两个形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-148">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="494ce-149">移动形状并调整其大小</span><span class="sxs-lookup"><span data-stu-id="494ce-149">Move and resize shapes</span></span>

<span data-ttu-id="494ce-150">形状位于工作表的顶部。</span><span class="sxs-lookup"><span data-stu-id="494ce-150">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="494ce-151">它们的位置由`left`和`top`属性定义。</span><span class="sxs-lookup"><span data-stu-id="494ce-151">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="494ce-152">这些操作充当工作表各自边缘的边距, [0, 0] 为左上角。</span><span class="sxs-lookup"><span data-stu-id="494ce-152">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="494ce-153">这些值既可以直接设置`incrementLeft` , 也可以使用 and `incrementTop`方法从当前位置进行调整。</span><span class="sxs-lookup"><span data-stu-id="494ce-153">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="494ce-154">从默认位置旋转的形状也是通过这种方式建立的, `rotation`属性是绝对量和调整现有旋转的`incrementRotation`方法。</span><span class="sxs-lookup"><span data-stu-id="494ce-154">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="494ce-155">相对于其他形状的形状的深度由该`zorderPosition`属性定义。</span><span class="sxs-lookup"><span data-stu-id="494ce-155">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="494ce-156">这是使用`setZOrder`方法进行设置的, 该方法采用[ShapeZOrder](/javascript/api/excel/excel.shapezorder)。</span><span class="sxs-lookup"><span data-stu-id="494ce-156">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="494ce-157">`setZOrder`调整当前形状相对于其他形状的排序。</span><span class="sxs-lookup"><span data-stu-id="494ce-157">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="494ce-158">您的外接程序有几个用于更改形状的高度和宽度的选项。</span><span class="sxs-lookup"><span data-stu-id="494ce-158">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="494ce-159">设置`height`或`width`属性将更改指定的维度, 而不更改其他维度。</span><span class="sxs-lookup"><span data-stu-id="494ce-159">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="494ce-160">和调整相对于当前或原始尺寸的形状各自的尺寸 (基于提供的 ShapeScaleType 的值)。 [](/javascript/api/excel/excel.shapescaletype) `scaleHeight` `scaleWidth`</span><span class="sxs-lookup"><span data-stu-id="494ce-160">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="494ce-161">可选的[ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom)参数指定形状的缩放位置 (左上角、中间或右下角)。</span><span class="sxs-lookup"><span data-stu-id="494ce-161">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="494ce-162">如果该`lockAspectRatio`属性为**true**, 则缩放方法还会通过调整其他尺寸来保持形状的当前纵横比。</span><span class="sxs-lookup"><span data-stu-id="494ce-162">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="494ce-163">对`height`和`width`属性的直接更改仅影响该属性, 而不考虑`lockAspectRatio`属性的值。</span><span class="sxs-lookup"><span data-stu-id="494ce-163">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="494ce-164">下面的代码示例显示了缩放到其原始大小为1.25 倍且旋转30度的形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-164">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="494ce-165">形状中的文本</span><span class="sxs-lookup"><span data-stu-id="494ce-165">Text in shapes</span></span>

<span data-ttu-id="494ce-166">几何形状可以包含文本。</span><span class="sxs-lookup"><span data-stu-id="494ce-166">Geometric Shapes can contain text.</span></span> <span data-ttu-id="494ce-167">形状具有类型`textFrame`为[TextFrame](/javascript/api/excel/excel.textframe)的属性。</span><span class="sxs-lookup"><span data-stu-id="494ce-167">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="494ce-168">`TextFrame`对象管理文本显示选项 (如边距和文本溢出)。</span><span class="sxs-lookup"><span data-stu-id="494ce-168">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="494ce-169">`TextFrame.textRange`是一个带有 "文本内容" 和 "字体" 设置的[TextRange](/javascript/api/excel/excel.textrange)对象。</span><span class="sxs-lookup"><span data-stu-id="494ce-169">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="494ce-170">下面的代码示例创建一个名为 "Wave" 的几何形状, 其中包含文本 "shape text"。</span><span class="sxs-lookup"><span data-stu-id="494ce-170">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="494ce-171">它还调整形状和文本颜色, 并将文本的水平对齐方式设置为居中。</span><span class="sxs-lookup"><span data-stu-id="494ce-171">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="494ce-172">创建具有白色背景和`GeometricShape`黑色文本`Rectangle`的类型的`addTextBox`方法。 `ShapeCollection`</span><span class="sxs-lookup"><span data-stu-id="494ce-172">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="494ce-173">这与 Excel 的 "**插入**" 选项卡上的 "**文本框**" 按钮所创建的内容`addTextBox`相同。采用字符串参数设置的文本`TextRange`。</span><span class="sxs-lookup"><span data-stu-id="494ce-173">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="494ce-174">下面的代码示例演示如何创建带有文本 "Hello!" 的文本框。</span><span class="sxs-lookup"><span data-stu-id="494ce-174">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="494ce-175">形状组</span><span class="sxs-lookup"><span data-stu-id="494ce-175">Shape groups</span></span>

<span data-ttu-id="494ce-176">可以将形状组合在一起。</span><span class="sxs-lookup"><span data-stu-id="494ce-176">Shapes can be grouped together.</span></span> <span data-ttu-id="494ce-177">这样一来, 用户可以将它们视为定位、调整大小和其他相关任务的单个实体。</span><span class="sxs-lookup"><span data-stu-id="494ce-177">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="494ce-178">[ShapeGroup](/javascript/api/excel/excel.shapegroup)是一种类型的`Shape`, 因此加载项将该组视为单个形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-178">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="494ce-179">下面的代码示例演示组合在一起的三个形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-179">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="494ce-180">后续代码示例显示了形状组被移至右侧50像素。</span><span class="sxs-lookup"><span data-stu-id="494ce-180">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="494ce-181">组中的单个形状是通过`ShapeGroup.shapes`属性引用的, 该属性的类型为[GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection)。</span><span class="sxs-lookup"><span data-stu-id="494ce-181">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="494ce-182">分组后, 将无法再通过工作表的形状集合访问它们。</span><span class="sxs-lookup"><span data-stu-id="494ce-182">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="494ce-183">例如, 如果您的工作表中有三个形状, 并且它们都组合在一起, `shapes.getCount`则工作表的方法将返回一个计数为1。</span><span class="sxs-lookup"><span data-stu-id="494ce-183">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="494ce-184">将形状导出为图像</span><span class="sxs-lookup"><span data-stu-id="494ce-184">Export shapes as images</span></span>

<span data-ttu-id="494ce-185">任何`Shape`对象都可以转换为图像。</span><span class="sxs-lookup"><span data-stu-id="494ce-185">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="494ce-186">[getAsImage](/javascript/api/excel/excel.shape#getasimage-format-)返回 base64 编码的字符串。</span><span class="sxs-lookup"><span data-stu-id="494ce-186">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="494ce-187">图像的格式被指定为传递给[](/javascript/api/excel/excel.pictureformat) `getAsImage`的 PictureFormat 枚举。</span><span class="sxs-lookup"><span data-stu-id="494ce-187">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="494ce-188">删除形状</span><span class="sxs-lookup"><span data-stu-id="494ce-188">Delete shapes</span></span>

<span data-ttu-id="494ce-189">使用`Shape`对象的`delete`方法从工作表中删除形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-189">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="494ce-190">无需任何其他元数据。</span><span class="sxs-lookup"><span data-stu-id="494ce-190">No other metadata is needed.</span></span>

<span data-ttu-id="494ce-191">下面的代码示例从**MyWorksheet**中删除所有形状。</span><span class="sxs-lookup"><span data-stu-id="494ce-191">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="494ce-192">另请参阅</span><span class="sxs-lookup"><span data-stu-id="494ce-192">See also</span></span>

- [<span data-ttu-id="494ce-193">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="494ce-193">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="494ce-194">使用 Excel JavaScript API 处理图表</span><span class="sxs-lookup"><span data-stu-id="494ce-194">Work with Charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
