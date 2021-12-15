---
title: PowerPoint JavaScript 预览 API
description: 有关即将推出的 JavaScript PowerPoint的详细信息。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 406808b4b4ff16df72d9c37468696525c8be642f
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513989"
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

下表列出了当前预览PowerPoint JavaScript API 的列表。 有关所有 JavaScript POWERPOINT的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API。](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|指定段落中的项目符号是否可见。|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|表示段落的项目符号格式。|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|表示段落的水平对齐方式。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|返回此形状的填充格式。|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|指定形状的高度（以点表示）。|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|从形状左侧到幻灯片左侧的距离（以点表示）。|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|返回此形状的线条格式。|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|指定此形状的名称。|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|返回此形状的文本框对象。|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|从形状的上边缘到幻灯片上边缘的距离（以点表示）。|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|返回此形状的类型。|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|指定形状的宽度（以点表示）。|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|指定形状的高度（以点表示）。|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|指定从形状左侧到幻灯片左侧的距离（以点表示）。|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|指定从形状的上边缘到幻灯片上边缘的距离（以点表示）。|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|指定形状的宽度（以点表示）。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape (geometricShapeType：PowerPoint。GeometricShapeType，options？：PowerPoint。ShapeAddOptions) ](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|向幻灯片添加几何形状。|
||[addLine (connectorType？： PowerPoint。ConnectorType，选项？：PowerPoint。ShapeAddOptions) ](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|向幻灯片添加一行。|
||[addTextBox (text： string， options？： PowerPoint。ShapeAddOptions) ](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|向幻灯片添加一个文本框，并将提供的文本作为内容。|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|以 HTML 颜色格式表示形状填充前景色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|将形状的填充格式设置为统一颜色。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|将填充的透明度百分比指定为从 0.0 到 1.0 (不透明) 1.0 (透明) 。|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|返回形状的填充类型。|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|文本颜色格式的 HTML 颜色代码表示 (例如，"#FF0000"表示红色) 。|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|表示字体的斜体状态。|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|表示字体名称 (例如"Calibri") 。|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|表示字号（以 (，例如 11) ）。|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|应用于字体的下划线类型。|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|表示 HTML 颜色格式的线条颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|表示线条的虚线样式。|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|表示形状的线条样式。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|将线条的透明度百分比指定为从 0.0 到 1.0 (不透明) 1.0 (透明) 。|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|指定形状元素的线条格式是否可见。|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|表示线条的粗细（以磅为单位）。|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|文本框的自动大小调整设置。|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|删除文本框中的所有文本。|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|指定文本框是否包含文本。|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|表示文本框的左边距（以磅为单位）。|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|表示文本框的右边距（以磅为单位）。|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|表示文本框的垂直对齐方式。|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|确定是否自动中断行以适合形状中的文本。|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|返回 `ShapeFont` 一个对象，该对象代表文本范围的字体属性。|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|返回 `TextRange` 给定范围中子字符串的对象。|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|代表文本范围的段落格式。|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|表示文本范围的纯文本内容。|

## <a name="see-also"></a>另请参阅

- [PowerPoint JavaScript API 参考文档](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md)