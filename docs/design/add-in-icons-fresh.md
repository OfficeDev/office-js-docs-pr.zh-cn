---
title: Office 外接程序的全新样式图标指南
description: Office 加载项中使用新鲜样式图标的指南。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: bd2cb372b79bef7f8c81deb778862f6bfd91d742
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467004"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Office 外接程序的全新样式图标指南

Office 2013+ (永久) 版本的 Office 使用 Microsoft 的 Fresh 样式图标。 如果希望图标与 Microsoft 365 的 Monoline 样式匹配，请参阅 [Office 外接程序的 Monoline 样式图标指南](add-in-icons-monoline.md)。

## <a name="office-fresh-visual-style"></a>Office 新鲜视觉样式

Fresh 图标仅包含重要的交互元素。 包括透视、渐变和光源的非必需元素均被删除。 简化后的图标可支持对命令和控件的快速解析。 遵循此样式以最适合 Office 永久客户端。

## <a name="best-practices"></a>最佳做法

创建图标时，请遵循以下准则。

|允许事项|禁止事项|
|:---|:---|
|保持视觉对象简单明了，重点关注通信的关键元素。| 不要使用使图标显得杂乱的项目。|
|使用 Office 图标语言来表示行为或概念。|不要在 Office 应用功能区或上下文菜单中将外接程序命令的 Fabric Core 标志符号重新设置为用途。 Fabric Core 图标在风格上不同，不匹配。|
|将画笔等公用 Office 视觉隐喻重用于格式或用于查找的放大镜。|不要对不同的命令重复使用视觉隐喻。 对不同的行为和概念使用同一图标可能会引起混淆。 |
|重绘图标，使其更大或更小。 请花时间重绘切割区、角和圆边，以最大化线条的清晰度。 |切勿通过缩小或扩大尺寸来调整图标大小。 这可能会导致视觉对象质量不佳和操作不清晰。 对于较大尺寸的复杂图标，如果不是通过重绘来使其变小，则可能会降低清晰度。 |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  |避免依赖徽标或品牌传达外接程序命令应起到的作用。 品牌标志在较小的图标尺寸上和应用很多修饰符后并非总具有识别性。 品牌标记通常与 Office 应用功能区图标样式冲突，并且可以在饱和环境中争夺用户的注意力。 |
|使用具有透明背景的 PNG 格式。 |*没有。*|
|避免在图标中使用可本地化的内容，包括印刷字符、段落标记指示和问号。 |*没有。*|

## <a name="icon-size-recommendations-and-requirements"></a>图标大小的建议和要求

Office 桌面图标是位图图像。 根据用户的 DPI 设置和触摸模式将呈现不同的大小。 包括所有八种支持的大小，可在所有受支持的解决方案和上下文中创建最佳体验。 下面是支持的大小 - 需要三个大小。

- 16 像素（必需）
- 20 像素
- 24 像素
- 32 像素（必需）
- 40 像素
- 48 像素
- 64 像素（建议，最适用于 Mac）
- 80 像素（必需）

> [!IMPORTANT]
> 有关作为外接程序代表图标的映像，请参阅 [AppSource 和 Office 中的有效列表](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) ，了解大小和其他要求。

确保根据每个尺寸重新绘制你的图标，而非将其缩小。

![建议按大小重绘图标，而不是收缩图标。 例如，可能需要在小型图标中使用更少的元素，而不只是向下缩放更大的图像。](../images/icon-resizing.png)

## <a name="icon-anatomy-and-layout"></a>图标分析和布局

Office 图标通常由包含操作和概念修饰符的基元素组成。 操作修饰符表示诸如添加、打开、新建或关闭等的概念。 概念修饰符表示图标的状态、更改或说明。

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

以下图像显示 Office 图标中的基本元素和修饰符的布局。

![显示中间的图标基元素，右下角有修饰符，左上角有操作修饰符。](../images/icon-layouts.png)

- 将基本元素置于像素框架的中间位置，并在其周围填充空白。
- 将操作修饰符置于左上方。
- 将概念修饰符置于右下方。
- 限制图标中的元素数。 在 32 px 时，将修饰符数限制为最多 2 个。 在 16 px 时，将修饰符数限制为 1。

### <a name="base-element-padding"></a>基准元素填充

放置与大小相一致的基本元素。 如果基本元素不能在框架居中显示，则将其对齐到左上方，并将多余的像素保留在右下方。 为了获得最佳结果，请应用下一部分中表中列出的填充指南。

### <a name="modifiers"></a>修饰符

所有修饰符应在每个元素（包括背景）之间有一个 1 px 透明切口。 元素不应直接重叠。 在规则和边缘之间创建空白空间。 修饰符在大小上可能略有不同，但会将这些尺寸作为起点使用。

|图标大小|在基本元素周围填充|修饰符大小|
|:---|:---|:---|
|16 px|0|9 px|
|20 像素|1px|10 像素|
|24 像素|1px|12 px|
|32 px|2px|14 像素|
|40 像素|2px|20 像素|
|48 像素|3px|22 px|
|64 px|5px|29 px|
|80 px|5px|38 px|

## <a name="icon-colors"></a>图标颜色

> [!NOTE]
> 这些颜色指南适用于[外接程序命令](add-in-commands.md)中使用的功能区图标。 这些图标不会使用 Fluent UI 呈现，调色板与 Microsoft UI Fabric |中描述的调色板不同[颜色|共享](https://fluentfabric.azurewebsites.net/#/color/shared)。

Office 图标具有一个有限的调色板。 使用下表中列出的颜色确保与 Office UI 无缝集成。 对颜色的使用应用以下准则。

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.
- 如有可能，除灰色外仅使用其他一种颜色。 将附加颜色限制为最多两种。
- 所有图标大小中的颜色应具有一致的外观。 Office 图标针对不同的图标大小具有略微不同的调色板。 16 个 px 和更小的图标比 32 px 和更大的图标略深，更有活力。 除了这些细微的调整以外，颜色的差别体现在大小上。

|颜色名称|RGB|十六进制|颜色|类别|
|:---|:---|:---|:---|:---|
|文本灰色 (80)|80、80、80|#505050| ![文本的灰色 80 颜色。](../images/color-text-gray-80.png) |文本|
|文本灰色 (95)|95、95、95|#5F5F5F| ![文本的灰色 95 颜色。](../images/color-text-gray-95.png) |文本|
|文本灰色 (105)|105、105、105|#696969| ![文本的灰色 105 颜色。](../images/color-text-gray-105.png) |文本|
|深灰色 32|128、128、128|#808080| ![32 px 及更大的深灰色。](../images/color-dark-gray-32.png) |32 px 及更高版本|
|中灰色 32|158、158、158|#9E9E9E| ![32 px 及更大的中灰色。](../images/color-medium-gray-32.png) |32 px 及更高版本|
|浅灰色所有|179、179、179|#B3B3B3| ![所有图像大小的浅灰色。](../images/color-light-gray-all.png) |所有大小|
|深灰色 16|114、114、114|#727272| ![16 px 及更小的深灰色。](../images/color-dark-gray-16.png) |16 px 及更低|
|中灰色 16|144、144、144|#909090| ![16 px 和更小的中灰色。](../images/color-medium-gray-16.png) |16 及以下|
|蓝色 32|77、130、184|#4d82B8| ![32 px 及更大的蓝色。](../images/color-blue-32.png) |32 px 及更高版本|
|蓝色 16|74、125、177|#4A7DB1| ![16 px 和更小的蓝色。](../images/color-blue-16.png) |16 px 及更低|
|黄色所有|234、194、130|#EAC282| ![所有图像大小的黄色。](../images/color-yellow-all.png) |所有大小|
|橙色 32|231、142、70|#E78E46| ![32 px 及更大的橙色。](../images/color-orange-32.png) |32 px 及更高版本|
|橙色 16|227、142、70|#E3751C| ![16 px 和更小的橙色。](../images/color-orange-16.png) |16 px 及更低|
|粉色所有|230、132、151|#E68497| ![所有图像大小的粉红色。](../images/color-pink-all.png) |所有大小|
|绿色 32|118、167、151|#76A797| ![32 px 及更大的绿色。](../images/color-green-32.png) |32 px 及更高版本|
|绿色 16|104、164、144|#68A490| ![16 px 及更小的绿色。](../images/color-green-16.png) |16 px 及更低|
|红色 32|216、99、68|#D86344| ![32 px 及更大的红色。](../images/color-red-32.png) |32 px 及更高版本|
|红色 16|214、85、50|#D65532| ![16 px 及更小的红色。](../images/color-red-16.png) |16 px 及更低|
|紫色 32|152、104、185|#9868B9| ![32 px 及更大的紫色。](../images/color-purple-32.png) |32 px 及更高版本|
|紫色 16|137、89、171|#8959AB| ![16 px 及更小的紫色。](../images/color-purple-16.png) |16 px 及更低|

## <a name="icons-in-high-contrast-modes"></a>高对比度模式下的图标

Office 图标设计为在高对比度模式中完美呈现。 前景元素与最大化易读性和启用重新着色的背景明显不同。 在高对比度模式下，Office 会使用小于 190 的红色、绿色或蓝色值直到全黑，为任何像素的图标重新着色。 其他所有像素都将是白色的。 换言之，每个评估的 RGB 通道中的 0-189 值表示为黑色，而 190-255 值表示为白色。 其他高对比度主题则使用相同的 190 阈值但不同的规则进行重新着色。 例如，高对比度白色主题会将所有大于 190 的像素重新着色为不透明，而将所有其他像素重新着色为透明。 应用以下准则以在高对比度设置中最大程度地提高可读性。

- 旨在以 190 阈值区分前景和背景元素。
- 遵循 Office 图标视觉样式。
- 使用图标调色板中的颜色。
- 避免使用渐变。
- 避免使用值相似的颜色块。

## <a name="see-also"></a>另请参阅

- [图标清单元素](/javascript/api/manifest/icon)
- [IconUrl manifest 元素](/javascript/api/manifest/iconurl)
- [HighResolutionIconUrl manifest 元素](/javascript/api/manifest/highresolutioniconurl)
- [创建加载项图标](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
