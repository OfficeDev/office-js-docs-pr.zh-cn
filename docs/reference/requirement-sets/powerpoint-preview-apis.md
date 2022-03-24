---
title: PowerPoint JavaScript 预览 API
description: 有关即将推出的 JavaScript PowerPoint的详细信息。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 2d43ca19d36b9f30e8699370bc97ecf194395d06
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742944"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint JavaScript 预览 API

JavaScript API PowerPoint在"预览"中首次引入，之后在经过充分测试并获取用户反馈后，它将成为特定编号要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 幻灯片管理 | 添加对添加幻灯片以及管理幻灯片版式和幻灯片母版的支持。 | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| 形状 | 添加对获取对幻灯片中形状的引用的支持。 | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>API 列表

下表列出了当前预览PowerPoint JavaScript API 的列表。 有关所有 JavaScript POWERPOINT的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#powerpoint-powerpoint-bulletformat-visible-member)|指定段落中的项目符号是否可见。|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-bulletformat-member)|表示段落的项目符号格式。|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-horizontalalignment-member)|表示段落的水平对齐方式。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-fill-member)|返回此形状的填充格式。|
||[height](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-height-member)|指定形状的高度（以点表示）。|
||[left](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-left-member)|从形状左侧到幻灯片左侧的距离（以点表示）。|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-lineformat-member)|返回此形状的线条格式。|
||[name](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-name-member)|指定此形状的名称。|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-textframe-member)|返回此形状的文本框对象。|
||[top](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-top-member)|从形状的上边缘到幻灯片上边缘的距离（以点表示）。|
||[type](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-type-member)|返回此形状的类型。|
||[width](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-width-member)|指定形状的宽度（以点表示）。|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-height-member)|指定形状的高度（以点表示）。|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-left-member)|指定从形状左侧到幻灯片左侧的距离（以点表示）。|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-top-member)|指定从形状的上边缘到幻灯片上边缘的距离（以点表示）。|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-width-member)|指定形状的宽度（以点表示）。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape (geometricShapeType：PowerPoint。GeometricShapeType， options？： PowerPoint.ShapeAddOptions) ](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgeometricshape-member(1))|向幻灯片添加几何形状。|
||[addLine (connectorType？：PowerPoint。ConnectorType，options？：PowerPoint。ShapeAddOptions) ](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addline-member(1))|向幻灯片添加一行。|
||[addTextBox (text： string， options？： PowerPoint。ShapeAddOptions) ](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1))|向幻灯片添加一个文本框，并将提供的文本作为内容。|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-clear-member(1))|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-foregroundcolor-member)|以 HTML 颜色格式表示形状填充前景色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-setsolidcolor-member(1))|将形状的填充格式设置为统一颜色。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-transparency-member)|将填充的透明度百分比指定为从 0.0 到 1.0 (不透明) 1.0 (透明) 。|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-type-member)|返回形状的填充类型。|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-bold-member)|表示字体的加粗状态。|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-color-member)|文本颜色格式的 HTML 颜色代码表示 (例如，"#FF0000"表示红色) 。|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-italic-member)|表示字体的斜体状态。|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-name-member)|表示字体名称 (例如"Calibri") 。|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-size-member)|表示字号（以 (，例如 11) ）。|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-underline-member)|应用于字体的下划线类型。|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-color-member)|表示 HTML 颜色格式的线条颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-dashstyle-member)|表示线条的虚线样式。|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-style-member)|表示形状的线条样式。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-transparency-member)|将线条的透明度百分比指定为从 0.0 到 1.0 (不透明) 1.0 (透明) 。|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-visible-member)|指定形状元素的线条格式是否可见。|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-weight-member)|表示线条的粗细（以磅为单位）。|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-autosizesetting-member)|文本框的自动大小调整设置。|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-bottommargin-member)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-deletetext-member(1))|删除文本框中的所有文本。|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-hastext-member)|指定文本框是否包含文本。|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-leftmargin-member)|表示文本框的左边距（以磅为单位）。|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-rightmargin-member)|表示文本框的右边距（以磅为单位）。|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-textrange-member)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-topmargin-member)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-verticalalignment-member)|表示文本框的垂直对齐方式。|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-wordwrap-member)|确定是否自动中断行以适合形状中的文本。|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-font-member)|`ShapeFont`返回一个对象，该对象代表文本范围的字体属性。|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-getsubstring-member(1))|`TextRange`返回给定范围中子字符串的对象。|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-paragraphformat-member)|代表文本范围的段落格式。|
||[text](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-text-member)|表示文本范围的纯文本内容。|

## <a name="see-also"></a>另请参阅

- [PowerPoint JavaScript API 参考文档](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md)