---
title: 适用于 Office 外接程序的 Monoline 样式图标准则
description: 获取有关在 Office 外接程序中使用 Monoline 样式图标图标的指南。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 264aa9e01bd70924cfee01a864c515c8c7a4d138
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132198"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>适用于 Office 外接程序的 Monoline 样式图标准则

Monoline style 插图在 Office 365 中使用。 如果您希望图标与非订阅 Office 2013 + 的新样式相匹配，请参阅 [Office 外接程序的新样式图标指南](add-in-icons-fresh.md)。

## <a name="office-monoline-visual-style"></a>Office Monoline 视觉样式

Monoline 样式的目标具有一致、清楚和可访问的插图若要通过简单的视觉对象传达操作和功能，请确保所有用户都可以访问图标，并且具有与在 Windows 中其他地方使用的样式一致的样式。

以下准则适用于要为其创建与已存在 Office 产品的图标一致的功能的图标的第三方开发人员。

### <a name="design-principles"></a>设计原则

- 简单、干净、清晰。
- 仅包含必要的元素。
- 受 Windows 图标样式的灵感。
- 所有用户均可访问。

#### <a name="conveying-meaning"></a>传达含义

- 使用描述性元素（如页面）表示文档或表示邮件的信封。
- 使用相同的元素表示相同的概念，即邮件始终由信封而不是图章表示。
- 在概念开发过程中使用核心比喻。

#### <a name="reduction-of-elements"></a>减小元素

- 将图标缩小为其核心含义，仅使用对隐喻至关重要的元素。
- 将图标中的元素数限制为两个，而不考虑图标大小。

#### <a name="consistency"></a>稳定性

图标的大小、排列和颜色应一致。

#### <a name="styling"></a>样式

##### <a name="perspective"></a>Perspective

默认情况下，Monoline 图标是面向前的。 允许使用透视和/或旋转的某些元素（如多维数据集），但应将异常保持为最小值。

##### <a name="embellishment"></a>Embellishment

Monoline 是一个简洁的最小样式。 所有内容都使用单色，这意味着没有渐变、纹理或光源。

## <a name="designing"></a>设计

### <a name="sizes"></a>大小

我们建议您在所有这些尺寸中生成每个图标以支持高 DPI 设备。 绝对 *必需* 的大小为 16 px、20 px 和32像素，因为它们是100% 的大小。

**16 px、20 px、24 px、32 px、40 px、48 px、64 px、80 px、96 px**

### <a name="layout"></a>布局

下面是带有修饰符的图标布局的示例。

![右下角带有修饰的图标的关系图](../images/monolineicon1.png)  ![具有相同图标的关系图，其中添加了网格背景和用于基、修饰符、填充和剪切的标注](../images/monolineicon2.png)

#### <a name="elements"></a>元素

- **基本**：图标表示的主要概念。 这通常是图标的唯一需要的视觉对象，但有时主要概念可以通过辅助元素（修饰符）进行增强。

- **修饰符** 覆盖基准的任何元素;即，通常表示操作或状态的修饰符。 它通过充当添加、改变或描述符来修改 base 元素。

![具有被称为 "base" 和 "修改区域" 的网格关系图](../images/monolineicon3.png)

### <a name="construction"></a>建造

#### <a name="element-placement"></a>元素放置

将 Base 元素放置在填充内图标的中心。 如果不能完全居中放置，则 base 应为错误的右上。 在下面的示例中，该图标是完全居中的。

![显示完全居中图标的图示](../images/monolineicon4.png)

在下面的示例中，向左 erring 图标。

![显示 errs 左侧为1像素的图标的图示](../images/monolineicon5.png)

修饰符几乎总是放置在图标画布的右下角。 在极少数情况下，修饰符放置在不同的角。 例如，如果 base 元素不可识别的右下角的修饰符，请考虑将其放在左上角。

![显示四个图标的图，其中右下侧有一个修改框，在左上角显示一个带有修饰符的图标](../images/monolineicon6.png)

#### <a name="padding"></a>填充

每个大小图标在图标周围都有指定的填充量。 Base 元素保持在填充范围内，但该修饰符应对接到画布的边缘，延伸到图标边框的边缘之外。 下面的图像显示了建议用于每个图标大小的填充。

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![带有0px 填充的 16 px 图标](../images/monolineicon7.png)|![具有1px 填充的 20 px 图标](../images/monolineicon8.png)|![带有1px 填充的 24 px 图标](../images/monolineicon9.png)|![使用2px 填充的 32 px 图标](../images/monolineicon10.png)|![使用2px 填充的 40 px 图标](../images/monolineicon11.png)|![使用3px 填充的 48 px 图标](../images/monolineicon12.png)|![使用4px 填充的 64 px 图标](../images/monolineicon13.png)|![使用5px 填充的 80 px 图标](../images/monolineicon14.png)|![使用6px 填充的 96 px 图标](../images/monolineicon15.png)|

#### <a name="line-weights"></a>线条粗细

Monoline 是一种由直线和分级显示的形状所占据的样式。 根据生成图标的大小，应使用以下线条粗细。

|图标大小：|16px|20px|24px|32px|40px|48px|64px|80px|96px|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**线条粗细：**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
|**示例图标：**|![16 px 图标](../images/monolineicon16.png)|![20 px 图标](../images/monolineicon17.png)|![24 px 图标](../images/monolineicon18.png)|![32 px 图标](../images/monolineicon19.png)|![40 px 图标](../images/monolineicon20.png)|![48 px 图标](../images/monolineicon21.png)|![64 px 图标](../images/monolineicon22.png)|![80 px 图标](../images/monolineicon23.png)|![96 px 图标](../images/monolineicon24.png)|

#### <a name="cutouts"></a>块

如果将 icon 元素放置在另一个元素的顶部，将使用底部元素) 的切除 (来提供这两个元素之间的空间，这主要是出于可读性目的。 在 base 元素的顶部放置修饰符时，通常会发生这种情况，但在某些情况下，这两个元素都不是修饰符。 这两个元素之间的这两种切口有时称为 "间隙"。

间隙大小的宽度应与用于该大小的线条粗细的宽度相同。 如果使用16像素的图标，则间隙宽度将为1px，如果为 48 px 图标，则间隙应为2px。 下面的示例显示一个 32 px 图标，在修饰符和基础基之间存在间隔1px。

![在修饰符和基础基底之间存在间隔为1px 的 32 px 图标](../images/monolineicon25.png)

在某些情况下，如果修饰符具有对角线或曲线边缘且标准间隙不能提供足够的分隔，则间隙可能会增加1/2 像素。 这可能只会影响具有1px 线宽的图标： 16 px、20 px、24 px 和32像素。

#### <a name="background-fills"></a>背景填充

Monoline 图标集中的大多数图标都需要进行背景填充。 但是，在某些情况下，对象不会发生自然填充，因此不应应用填充。 以下图标具有白色填充。

![使用白填充的五个图标的编译](../images/monolineicon26.png)

以下图标没有填充。  (包括齿轮图标以显示未填充中心孔。 ) 

![无填充的五个图标的编译](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>填充的最佳实践

###### <a name="dos"></a>进入

- 填充具有定义的边界且自然具有填充的任何元素。
- 使用单独的形状创建背景填充。
- 使用 [调色板](#color)中的 "**背景填充**"。
- 维护重叠元素之间的像素分隔。
- 在多个对象之间进行填充。

###### <a name="donts"></a>应该

- 不填充不自然填充的对象;例如，曲别针。
- 不填充方括号。
- 不要在数字或字母字符后面填写。

### <a name="color"></a>颜色

调色板设计为简单性和可访问性。 它包含4种中性色和两种蓝色、绿色、黄色、红色和紫色的变体。 "Monoline" 图标颜色调色板中不会有意包含橙色。 每种颜色都旨在以特定方式使用，如本节中所述。

#### <a name="palette"></a>调色板

![Monoline 中的四个灰色阴影：独立或轮廓为深灰色; 对于大纲或内容为浅灰色; 对于背景填充为浅灰色，填充浅灰色](../images/monoline-grayshades.png)

![Monoline 中的调色板包括一种蓝色、绿色、黄色、红色和紫色的底纹，用于独立、轮廓和填充](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>如何使用颜色

在 Monoline 调色板中，所有颜色都具有独立、轮廓和填充的变体。 通常情况下，使用填充和边框构造元素。 颜色以下列模式之一应用：

- 独立于没有填充的对象的单独颜色。
- 边框使用边框颜色，填充使用填充颜色。
- 边框使用独立的颜色，填充使用背景填充颜色。

以下是使用颜色的示例。

![在边框或填充中的颜色和/或填充的三个图标的编译](../images/monolineicon28.png)

最常见的情况是让元素将深灰色独立用于背景填充。

使用彩色填充时，应始终使用其对应的边框颜色。 例如，蓝色填充应仅与蓝色轮廓一起使用。 但是，此常规规则有两个例外：

- 背景填充可用于任何颜色独立。
- 浅灰色填充可用于两种不同的轮廓颜色：深灰色或中等灰度。

#### <a name="when-to-use-color"></a>何时使用颜色

应使用 Color 来传达图标的含义，而不是 embellishment 的含义。 它应将 **操作突出显示** 给用户。 将一个修饰符添加到具有颜色的 base 元素中时，基本元素通常会转换为深灰色和背景填充，以便修饰符可以是颜色的元素，如下面的示例将 "X" 修饰符添加到以下集合中最左边的图标中的图片基。

![使用颜色的五个图标的编译](../images/monolineicon29.png)

您应将图标限制为另 **一种** 颜色，而不是上面提到的轮廓和填充。 但是，如果对其比喻至关重要，则可以使用更多的颜色，而不是灰色之外的其他两种颜色限制。 在极少数情况下，如果需要更多颜色，也会出现异常。 以下是仅使用一种颜色的较简单的图标示例。

  ![每个使用一种颜色的五个图标的编译](../images/monolineicon30.png)

但以下图标使用的颜色过多。

  ![对每个使用多种颜色的五个图标进行编译](../images/monolineicon31.png)

对内部 "content" 使用 **中灰色** ，例如电子表格图标中的网格线。 当内容需要显示控件的行为时，将使用其他内部颜色。

![使用中等灰色的内部元素编译五个图标](../images/monolineicon32.png)

#### <a name="text-lines"></a>文本行

当文本行位于 "容器" 中时 (例如，文档中的文本) 上，使用 "中灰色"。 不在容器中的文本行应为 **深灰色**。

### <a name="text"></a>文本

避免在图标中使用文本字符。 由于世界各地使用的是 Office 产品，因此我们希望尽可能将图标保留为中性语言。

## <a name="production"></a>生产

### <a name="icon-file-format"></a>图标文件格式

最终图标应另存为 .png 图像文件。 使用具有透明背景且具有32位深度的 PNG 格式。
