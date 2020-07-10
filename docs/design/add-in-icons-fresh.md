---
title: Office 外接程序的全新样式图标准则
description: 获取有关在 Office 外接程序中使用全新样式图标图标的指南。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 7f29de70712448e9ee7458db864fb40746412153
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093929"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Office 外接程序的全新样式图标准则

Office 2013 + (非订阅) 版本的 Office 使用 Microsoft 的插图的全新样式。 如果您希望图标符合 Microsoft 365 的 Monoline 样式，请参阅[Office 加载项的 Monoline 样式图标准则](add-in-icons-monoline.md)。

## <a name="office-fresh-visual-style"></a>Office 全新视觉样式

新图标仅包含基本的 communicative 元素。 包括透视、渐变和光源的非必需元素均被删除。 简化后的图标可支持对命令和控件的快速解析。 遵循此样式以最适合 Office 非订阅客户端。

## <a name="best-practices"></a>最佳实践

创建图标时，请遵循以下准则：

|允许事项|禁止事项|
|:---|:---|
|保持可视化效果简单明了，重点关注通信的关键元素。| 不要使用使图标显得杂乱的项目。|
|使用 Office 图标语言来表示行为或概念。|请勿在 Office 应用程序功能区或上下文菜单中重新调整加载项命令的 Office UI Fabric 标志符号。 Fabric 图标风格不同，不能匹配。|
|将画笔等公用 Office 视觉隐喻重用于格式或用于查找的放大镜。|不要对不同的命令重复使用视觉隐喻。 对不同的行为和概念使用同一图标可能会引起混淆。 |
|重绘图标，使其更大或更小。 请花时间重绘切割区、角和圆边，以最大化线条的清晰度。 |切勿通过缩小或扩大尺寸来调整图标大小。 这可能会导致视觉对象质量不佳和操作不清晰。 对于较大尺寸的复杂图标，如果不是通过重绘来使其变小，则可能会降低清晰度。 |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  |避免依赖徽标或品牌传达外接程序命令应起到的作用。 品牌标志在较小的图标尺寸上和应用很多修饰符后并非总具有识别性。 品牌标记通常与 Office 应用程序功能区图标样式发生冲突，并且可以在饱和环境中争夺用户的注意力。 |
|使用具有透明背景的 PNG 格式。 ||
|避免在图标中使用可本地化的内容，包括印刷字符、段落标记指示和问号。 ||

## <a name="icon-size-recommendations-and-requirements"></a>图标大小的建议和要求

Office 桌面图标是位图图像。 根据用户的 DPI 设置和触摸模式将呈现不同的大小。 包括所有八种支持的大小，可在所有受支持的解决方案和上下文中创建最佳体验。 以下是受支持的大小 - 三种是必需的：

- 16 像素（必需）
- 20 像素
- 24 像素
- 32 像素（必需）
- 40 像素
- 48 像素
- 64 像素（建议，最适用于 Mac）
- 80 像素（必需）

确保根据每个尺寸重新绘制你的图标，而非将其缩小。

![显示调整图标大小而非缩小图标的建议的图示](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>图标分析和布局

Office icons are typically comprised of a base element with action and conceptual modifiers overlayed. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon.

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

以下图像显示 Office 图标中的基本元素和修饰符的布局。

![显示处于中间位置的图标基本元素的图像，其中修饰符位于右下方，操作修饰符位于左上方](../images/icon-layouts.png)

- 将基本元素置于像素框架的中间位置，并在其周围填充空白。
- 将操作修饰符置于左上方。
- 将概念修饰符置于右下方。
- Limit the number of elements in your icons. At 32px, limit the number of modifiers to a maximum of two. At 16px, limit the number of modifiers to one.

### <a name="base-element-padding"></a>基准元素填充

放置与大小相一致的基本元素。 如果基本元素不能在框架居中显示，则将其对齐到左上方，并将多余的像素保留在右下方。 为获得最佳结果，请应用下一节的表中列出的填充准则。

### <a name="modifiers"></a>修饰符

All modifiers should have a 1px transparent cutout between each element, including the background. Elements should not directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.

|**图标大小**|**在基本元素周围填充**|**修饰符大小**|
|:---|:---|:---|
|16px|0|9px|
|20px|1px|10px|
|24px|1px|12px|
|32px|2px|14px|
|40px|2px|20px|
|48px|3px|22px|
|64px|5px|29px|
|80px|5px|38px|

## <a name="icon-colors"></a>图标颜色

> [!NOTE]
> 这些颜色指南适用于[外接程序命令](add-in-commands.md)中使用的功能区图标。 这些图标不使用 Microsoft UI Fabric 呈现，调色板与 [Microsoft UI Fabric | 颜色 | 共享](https://fluentfabric.azurewebsites.net/#/color/shared)中描述的调色板不同。

Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color:

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark. 
- If possible, use only one additional color beyond gray. Limit additional colors to two at the most.
- Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16px and smaller icons are slightly darker and more vibrant than 32px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.

|**颜色名称**|**RGB**|**十六进制**|**颜色**|**类别**|
|:---|:---|:---|:---|:---|
|文本灰色 (80)|80、80、80|#505050| ![文本灰色 80 彩色图像](../images/color-text-gray-80.png) |文本|
|文本灰色 (95)|95、95、95|#5F5F5F| ![文本灰色 95 彩色图像](../images/color-text-gray-95.png) |文本|
|文本灰色 (105)|105、105、105|#696969| ![文本灰色 105 彩色图像](../images/color-text-gray-105.png) |文本|
|深灰色 32|128、128、128|#808080| ![深灰色 32 彩色图像](../images/color-dark-gray-32.png) |32 及以上|
|中灰色 32|158、158、158|#9E9E9E| ![中灰色 32 彩色图像](../images/color-medium-gray-32.png) |32 及以上|
|浅灰色所有|179、179、179|#B3B3B3| ![浅灰色所有彩色图像](../images/color-light-gray-all.png) |所有大小|
|深灰色 16|114、114、114|#727272| ![深灰色 16 彩色图像](../images/color-dark-gray-16.png) |16 及以下|
|中灰色 16|144、144、144|#909090| ![中灰色 16 彩色图像](../images/color-medium-gray-16.png) |16 及以下|
|蓝色 32|77、130、184|#4d82B8| ![蓝色 32 彩色图像](../images/color-blue-32.png) |32 及以上|
|蓝色 16|74、125、177|#4A7DB1| ![蓝色 16 彩色图像](../images/color-blue-16.png) |16 及以下|
|黄色所有|234、194、130|#EAC282| ![黄色所有彩色图像](../images/color-yellow-all.png) |所有大小|
|橙色 32|231、142、70|#E78E46| ![橙色 32 彩色图像](../images/color-orange-32.png) |32 及以上|
|橙色 16|227、142、70|#E3751C| ![橙色 16 彩色图像](../images/color-orange-16.png) |16 及以下|
|粉色所有|230、132、151|#E68497| ![粉色所有彩色图像](../images/color-pink-all.png) |所有大小|
|绿色 32|118、167、151|#76A797| ![绿色 32 彩色图像](../images/color-green-32.png) |32 及以上|
|绿色 16|104、164、144|#68A490| ![绿色 16 彩色图像](../images/color-green-16.png) |16 及以下|
|红色 32|216、99、68|#D86344| ![红色 32 彩色图像](../images/color-red-32.png) |32 及以上|
|红色 16|214、85、50|#D65532| ![红色 16 彩色图像](../images/color-red-16.png) |16 及以下|
|紫色 32|152、104、185|#9868B9| ![紫色 32 彩色图像](../images/color-purple-32.png) |32 及以上|
|紫色 16|137、89、171|#8959AB| ![紫色 16 彩色图像](../images/color-purple-16.png) |16 及以下|

## <a name="icons-in-high-contrast-modes"></a>高对比度模式下的图标

Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings:

- 旨在以 190 阈值区分前景和背景元素。
- 遵循 Office 图标视觉样式。
- 使用图标调色板中的颜色。
- 避免使用渐变。
- 避免使用值相似的颜色块。
