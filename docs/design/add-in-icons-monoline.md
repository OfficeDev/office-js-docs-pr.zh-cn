---
title: Office 加载项的单行样式图标准则
description: 有关在 Office 加载项中使用 Monoline 样式图标的指南。
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7af7cbb7539ee2ae27efcadd4739f926cc81547a
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467067"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Office 加载项的单行样式图标准则

单行样式图标在 Office 应用中使用。 如果希望图标与永久 Office 2013+ 的“新鲜”样式匹配，请参阅 [Office 外接程序的“新鲜样式”图标指南](add-in-icons-fresh.md)。

## <a name="office-monoline-visual-style"></a>Office 单线视觉样式

Monoline 样式的目标是具有一致、清晰且易于访问的图标，以使用简单的视觉对象传达操作和功能，确保所有用户都可以访问图标，并且其样式与 Windows 中其他位置使用的图标一致。

以下准则适用于想要为与已存在的 Office 产品中的图标一致的功能创建图标的第三方开发人员。

### <a name="design-principles"></a>设计原则

- 简单、干净、清晰。
- 仅包含必要的元素。
- 灵感来自 Windows 图标样式。
- 可供所有用户访问。

#### <a name="convey-meaning"></a>传达含义

- 使用描述性元素（例如页面）来表示文档或信封来表示邮件。
- 使用同一元素表示相同的概念，即邮件始终由信封而不是邮票表示。
- 在概念开发过程中使用核心隐喻。

#### <a name="reduction-of-elements"></a>减少元素

- 将图标缩减为其核心含义，仅使用对隐喻至关重要的元素。
- 无论图标大小如何，将图标中的元素数限制为两个。

#### <a name="consistency"></a>一致性

图标的大小、排列和颜色应一致。

#### <a name="styling"></a>造型

##### <a name="perspective"></a>Perspective

默认情况下，单线图标面向前向。 允许某些需要透视和/或旋转的元素（如多维数据集），但应将异常保持在最低限度。

##### <a name="embellishment"></a>点缀

单线是一种干净的最小样式。 一切使用平面颜色，这意味着没有渐变、纹理或光源。

## <a name="designing"></a>设计

### <a name="sizes"></a>大小

建议在所有这些大小中生成每个图标，以支持高 DPI 设备。 绝对 *必需* 的大小为 16 px、20 px 和 32 px，因为这些大小是 100% 大小。

**16 px， 20 px， 24 px， 32 px， 40 px， 48 px， 64 px， 80 px， 96 px**

> [!IMPORTANT]
> 有关作为外接程序代表图标的映像，请参阅 [AppSource 和 Office 中的有效列表](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) ，了解大小和其他要求。

### <a name="layout"></a>布局

下面是带有修饰符的图标布局示例。

![右下方带有修饰符的图标图。](../images/monolineicon1.png)  ![同一图标，其中添加了基、修饰符、填充和切口的网格背景和标注。](../images/monolineicon2.png)

#### <a name="elements"></a>元素

- **基**：图标表示的主要概念。 这通常是图标所需的唯一视觉对象，但有时可以通过辅助元素（修饰符）增强主概念。

- **改 性 剂** 覆盖基的所有元素;也就是说，一个通常表示操作或状态的修饰符。 它通过充当加法、更改或描述符来修改基元素。

![标注了基本区域和修饰符区域的网格图。](../images/monolineicon3.png)

### <a name="construction"></a>建造

#### <a name="element-placement"></a>元素放置

基本元素放置在填充图标的中心。 如果不能完全居中，则基础应向右上偏右。 在下面的示例中，图标完全居中。

![显示完全居中图标的图示。](../images/monolineicon4.png)

在下面的示例中，图标在左侧出现问题。

![显示左侧 1 px 的图标。](../images/monolineicon5.png)

修饰符几乎总是放置在图标画布的右下角。 在一些极少数情况下，修饰符放置在其他角落。 例如，如果基元素在右下角的修饰符无法识别，请考虑将其放在左上角。

![显示四个图标，右下角有修饰符，一个图标左上角有修饰符。](../images/monolineicon6.png)

#### <a name="padding"></a>填充

每个大小图标在图标周围都有指定的填充量。 基元素保留在填充中，但修饰符应直至画布的边缘，在填充外部扩展到图标边框的边缘。 下图显示了要用于每个图标大小的建议填充。

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![带 0px 填充的 16 px 图标。](../images/monolineicon7.png)|![带有 1px 填充的 20 px 图标。](../images/monolineicon8.png)|![带有 1px 填充的 24 px 图标。](../images/monolineicon9.png)|![带 2px 填充的 32 px 图标。](../images/monolineicon10.png)|![带 2px 填充的 40 px 图标。](../images/monolineicon11.png)|![带 3px 填充的 48 px 图标。](../images/monolineicon12.png)|![带有 4px 填充的 64 px 图标。](../images/monolineicon13.png)|![带 5px 填充的 80 px 图标。](../images/monolineicon14.png)|![带有 6px 填充的 96 px 图标。](../images/monolineicon15.png)|

#### <a name="line-weights"></a>线条权重

单线是一种以线条和轮廓形状为主的样式。 根据生成图标的大小，应使用以下行权重。

|图标大小：|16px|20px|24px|32px|40px|48px|64px|80px|96px|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**线条重量：**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
|**示例图标：**|![16 px 图标。](../images/monolineicon16.png)|![20 px 图标。](../images/monolineicon17.png)|![24 px 图标。](../images/monolineicon18.png)|![32 px 图标。](../images/monolineicon19.png)|![40 px 图标。](../images/monolineicon20.png)|![48 px 图标。](../images/monolineicon21.png)|![64 px 图标。](../images/monolineicon22.png)|![80 px 图标。](../images/monolineicon23.png)|![96 px 图标。](../images/monolineicon24.png)|

#### <a name="cutouts"></a>切口

当图标元素放在另一个元素的顶部时，底部元素) 的切口 (用于在两个元素之间提供空间，主要用于可读性。 当修饰符放置在基元素之上时，通常会发生这种情况，但在某些情况下，这两个元素都不是修饰符。 这两个元素之间的这些切口有时称为“间隙”。

间隙的大小应与在该大小上使用的线条粗细相同。 如果创建 16 px 图标，则间隙宽度为 1px，如果是 48 px 图标，则间隙应为 2px。 下面的示例显示了一个 32 px 图标，修饰符和基础基数之间的间隙为 1px。

![32 px 图标，修饰符与基础基之间的间隙为 1px。](../images/monolineicon25.png)

在某些情况下，如果修饰符具有对角线或曲线边缘且标准间隙未提供足够的分离，则间隙可能会增加 1/2 px。 这可能仅影响具有 1px 线条权重的图标：16 px、20 px、24 px 和 32 px。

#### <a name="background-fills"></a>背景填充

Monoline 图标集中的大多数图标都需要背景填充。 但是，在某些情况下，对象不会自然而然地具有填充，因此不应应用填充。 以下图标具有白色填充。

![使用白色填充编译五个图标。](../images/monolineicon26.png)

以下图标没有填充。  (齿轮图标包含在内，以显示中心孔未填充。) 

![编译五个没有填充的图标。](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>填充的最佳做法

###### <a name="dos"></a>推荐做法

- 填充任何具有定义边界且自然具有填充的元素。
- 使用单独的形状创建背景填充。
- 使用 [调色板](#color)中的 **背景填充**。
- 保持重叠元素之间的像素分离。
- 在多个对象之间填充。

###### <a name="donts"></a>注意事项

- 不要填充不会自然填充的对象;例如，一个纸质禜。
- 不要填充括号。
- 不要在数字或 alpha 字符后面填充。

### <a name="color"></a>颜色

调色板旨在简化和辅助功能。 它包含 4 种中性颜色和蓝色、绿色、黄色、红色和紫色的两种变体。 橙色有意不包含在单线图标调色板中。 每种颜色都以特定的方式使用，如本部分所述。

#### <a name="palette"></a>调色板

![单线的四种灰色色调：独立或轮廓的深灰色、大纲或内容的中灰色、背景填充的浅灰色和填充的浅灰色。](../images/monoline-grayshades.png)

![单色调色板包括蓝色、绿色、黄色、红色和紫色的阴影，用于独立、轮廓和填充。](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>如何使用颜色

在单线调色板中，所有颜色都有独立、轮廓和填充变体。 通常，元素是使用填充和边框构造的。 颜色在以下模式之一中应用。

- 对于没有填充的对象，单独使用独立颜色。
- 边框使用大纲颜色，填充使用填充颜色。
- 边框使用独立颜色，填充使用背景填充颜色。

下面是使用颜色的示例。

![在边框或填充或填充中使用颜色编译三个图标。](../images/monolineicon28.png)

最常见的情况是让元素将深灰色独立与背景填充结合使用。

使用彩色填充时，它应始终与相应的大纲颜色一起使用。 例如，蓝色填充应仅与蓝色轮廓一起使用。 但是，此常规规则有两个例外。

- 背景填充可以与任何颜色独立一起使用。
- 浅灰色填充可以与两种不同的轮廓颜色一起使用：深灰色或中灰色。

#### <a name="when-to-use-color"></a>何时使用颜色

颜色应用于传达图标的含义，而不是用于点缀。 它应 **向用户突出显示操作** 。 将修饰符添加到具有颜色的基元素时，基本元素通常会变成深灰色和背景填充，以便修饰符可以是颜色元素，如下面的示例，将“X”修饰符添加到下一组最左侧图标的图片基中。

![使用颜色的五个图标的编译。](../images/monolineicon29.png)

应将图标限制为 **一种** 其他颜色，而上述大纲和填充除外。 但是，如果颜色对于它的隐喻至关重要，则可以使用更多的颜色，其限制是两种其他颜色，而不是灰色。 在极少数情况下，需要更多颜色时会出现异常。 下面是仅使用一种颜色的图标的良好示例。

  ![编译五个图标，每个图标使用一种颜色。](../images/monolineicon30.png)

但以下图标使用的颜色太多。

  ![编译五个图标，每个图标使用多种颜色。](../images/monolineicon31.png)

对内部“内容”使用 **中灰色** ，例如电子表格图标中的网格线。 当内容需要显示控件的行为时，将使用其他内部颜色。

![使用中灰色内部元素编译五个图标。](../images/monolineicon32.png)

#### <a name="text-lines"></a>文本行

例如，当文本行位于“容器” (中时，文档) 上的文本将使用中灰色。 容器中不包含的文本行应为 **深灰色**。

### <a name="text"></a>Text

避免在图标中使用文本字符。 由于 Office 产品在世界各地使用，因此我们希望将图标保持为语言中性。

## <a name="production"></a>生产

### <a name="icon-file-format"></a>图标文件格式

最终图标应保存为.png图像文件。 将 PNG 格式与透明背景配合使用，并具有 32 位深度。

## <a name="see-also"></a>另请参阅

- [图标清单元素](/javascript/api/manifest/icon)
- [IconUrl manifest 元素](/javascript/api/manifest/iconurl)
- [HighResolutionIconUrl manifest 元素](/javascript/api/manifest/highresolutioniconurl)
- [创建加载项图标](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
