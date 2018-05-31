---
title: Office 加载项设计语言
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7d19714fa14fb374bcd41aa744c08929c228c94f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437519"
---
# <a name="office-add-in-design-language"></a>Office 加载项设计语言

Office 设计语言是一种简单明了的视觉对象系统，它可确保体验的一致性。它包含一组用于定义 Office 接口的可视化元素，包括：

- 标准字样
- 公用调色板
- 一组版式大小和权重
- 图标准则
- 共享图标资源
- 动画定义
- 常见组件

[Office UI Fabric](https://dev.office.com/fabric) 是用于通过 Office 设计语言构建的官方前端框架。使用 Fabric 是可选的，但它是确保外接程序感觉像是 Office 的自然扩展的最快方法。利用 Fabric 来设计和构建补充 Office 的外接程序。

许多 Office 外接程序与先前存在的品牌相关联。你可以保留外接程序中的强大品牌及其视觉对象或组件语言。与 Office 集成的同时寻找保留自己的视觉对象语言的机会。寻找方法将 Office 颜色、版式、图标或其他样式元素置换为你自己品牌的元素。在插入客户熟悉的控件和组件时，寻找遵循通用外接程序布局或 UX 设计模式的方法。

在 Office 内插入基于主要品牌的 HTML 的 UI 会对客户产生不一致性。找到一个能够在 Office 中无缝整合的平衡点，同时与你的服务或父品牌保持明确一致。如果外接程序不适合 Office，通常是因为样式元素发生冲突。例如，版式过大和网格关闭、颜色对比度鲜明或太过强烈，或者相比 Office 动画过多且行为有差异。控件或组件的外观和行为与 Office 标准相差甚远。

## <a name="typography"></a>版式
Segoe 是 Office 的标准字样。在外接程序中使用 Segoe，以与 Office 任务窗格、对话框和内容对象保持一致。Office UI Fabric 允许你访问 Segoe。它在方便使用的 CSS 类中为 Segoe 的全型斜坡提供了许多不同的字体粗细和大小。并非所有 Office UI Fabric 大小和权重在 Office 外接程序中看上去都很理想。若要和谐适合或避免冲突，请考虑使用 Fabric 类型斜坡的一个子集。这是建议在 Office 外接程序中使用的 Fabric 的基类列表。

|示例 |类 |大小 |权重 |建议的用法 |
|------ |----- |---- |------ |----------------- |
|![特大文本图像](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 像素 | Segoe 光 |<ul><li>此类大于 Office 中的所有其他版式元素。请谨慎使用以避免超越可视化层次结构。</li><li>避免在有限空间中的长字符串上使用。</li><li>在使用此类的文本周围提供充足的空白空间。</li><li>常用于首次运行的信息、特大元素或其他操作调用。</li></ul> |
|![特大文本图像](../images/add-in-typeramp-title.png)|.ms-font-xl |21 像素 |Segoe 光 | <ul><li>此类匹配 Office 应用程序的任务窗格标题。</li><li>请谨慎使用以避免出现平面版式层次结构。</li><li>通常用作对话框、页面或内容标题等顶级元素。</li></ul> |
|![特大文本图像](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 像素 |Segoe 半光 | <ul><li>此类是标题下方的第一级元素。</li><li>常用作副标题、导航元素或组标头。</li><ul> |
|![特大文本图像](../images/add-in-typeramp-body.png)|.ms-font-m |14 像素 |Segoe 正常 |<ul><li>通常用作加载项中的正文文本。</li><ul>|
|![Hero 文本图像](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 像素 | Segoe 正常 |<ul><li>通常由行、标题或字段标签用于时间戳等二级或三级文本。</li><ul>|
|![Hero 文本图像](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 像素 |Segoe 半加重 |<ul><li>应极少使用类型渐变中的最小步长。它仅供不需要辨别的情况使用。</li><ul>|

> [!NOTE]
> 这些基类不包含文本颜色。请对白色背景上的大多数文本使用 Fabric 的“主中性色”。

## <a name="color"></a>颜色
颜色通常用于强调品牌并加强视觉层次。 它有助于识别界面，并在体验中为客户提供引导。 在Office内部，颜色被用于相同的目标，但它被有目的地和最低限度地应用。 在任何时候它都不应过度影响客户的内容。 即使每个Office应用程序都使用自己的主色调进行标记，它也会被谨慎使用。

Office UI Fabric 包含一组默认主题颜色。当 Fabric 作为组件应用于 Office 外接程序或应用于布局时，相同的目标适用。颜色应传达层次结构，有目的地指导客户操作而不会干扰内容。Fabric 主题颜色可以向整体界面引入新的个性色。此新的个性色可能会与 Office 应用程序品牌产生冲突并干扰层次结构。换句话说，Fabric 在外接程序内部使用时可能会向整体界面引入新的个性色。此新的个性色可能会分散用户注意力并干扰整个层次结构。寻找避免冲突和干扰的方法。使用中性个性色或覆盖 Fabric 主题颜色，以匹配 Office 应用程序品牌或你自己的品牌颜色。

Office 应用程序使客户能够通过应用 Office UI 主题个性化设置其界面。客户可以在四个 UI 主题中进行选择来改变背景样式以及 Word、PowerPoint、Excel 和 Office 套件中其他应用程序的按钮。若要使外接程序感觉像是 Office 的一个固有部分并响应个性化设置，请使用我们的主题 API。例如，任务窗格背景颜色在某些主题中切换到深灰色。我们的主题 API 允许你照做并调整前景文本，以确保[辅助功能](add-in-design-guidelines.md#accessibility-guidelines)。

> [!NOTE]
> - 对于邮件和任务窗格外接程序，请使用 [Context.officeTheme](https://dev.office.com/reference/add-ins/shared/office.context.officetheme) 元素匹配 Office 应用程序的主题。此 API 当前仅在 Office 2016 中可用。
> - 对于 PowerPoint 内容加载项，请参阅[在 PowerPoint 加载项中使用 Office 主题](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)。

将下列一般原则应用于颜色：

* 尽量少使用颜色来显示层次结构和强调品牌。
* 过度使用单个应用于交互式和非交互式元素的个性色可能会导致混乱。例如，避免将相同颜色用于导航菜单中的选定和未选定项。
* 避免与 Office 品牌应用颜色产生不必要的冲突。
* 使用自己的品牌颜色来生成与服务或公司的关联。
* 确保可以访问所有文本。确保前景文本与后台之间的对比度比为 4.5:1。
* 注意色盲群体。不要仅使用颜色指示交互性和层次结构。
* 请参阅[图标指南](design-icons.md)，详细了解如何使用 Office 图标调色板设计加载项命令图标。

## <a name="layout"></a>布局
嵌入到 Office 中的每个 HTML 容器都将有一个布局。这些布局是外接程序的主屏幕。你将在其中创建使客户能够启动操作、修改设置、查看、滚动或导航内容的体验。设计在屏幕中具有一致布局的外接程序，以确保体验的连续性。如果你有客户熟悉使用的现有网站，请考虑重新使用现有网页中的布局。对它们进行调整以协调适应 Office HTML 容器。

有关布局指南，请参阅[任务窗格](task-pane-add-ins.md)、[内容](content-add-ins.md)和[对话框](dialog-boxes.md)。若要详细了解如何将 Office UI Fabric 组件装配到通用布局和用户体验流，请参阅[用户体验设计模式模板](ux-design-patterns.md)。

请遵循下面的一般布局指南：

*   避免 HTML 容器上的边距过窄或过宽。20 像素是理想的默认值。
*   有意对齐元素。额外缩进和新对齐点应该有助于可视化层次结构。
*   Office 接口在 4 像素网格上。旨在使元素之间的填充保持在 4 的倍数。
*   界面过于拥挤可能导致混乱，并抑制触控交互的易用性。
*   在各个屏幕之间保持布局一致性。意外布局更改类似于视觉错误，这将导致对解决方案的信心和信任的缺失。
*   遵循公用的布局模式。约定可帮助用户了解如何使用界面。
*   避免冗余元素，如品牌或命令。
*   整合控件和视图，以避免需要过多地移动鼠标。
*   创建适应 HTML 容器宽度和响应体验。

## <a name="component-language"></a>组件语言

屏幕和布局由内容和组件组成。组件是帮助客户与软件或服务元素进行交互的控件。按钮、导航、徽章、警报和下拉列表是常见组件的所有示例，通常具有一致的样式和行为。

Office UI Fabric 呈现外观和行为类似于 Office 部件的组件。利用 Fabric 与 Office 无缝集成。如果外接程序有其自己预先存在的组件语言，则不需要为支持 Fabric 而放弃它。与 Office 集成的同时寻找保留该语言的机会。寻找置换出风格元素、删除冲突，或采用样式和行为以避免用户混淆的方法。

将下列一般原则应用于组件：

*   请勿在外接程序中复制 Office 功能区
*   避免创建与 Office 组件行为不同的菜单、按钮或其他组件。
*   使用我们建议用于外接程序的 [Office UI Fabric](office-ui-fabric.md) 组件。
*   将 [UX 设计模式模板](ux-design-patterns.md)用于常用的 Office UI 组件。

## <a name="icons"></a>图标
图标是行为或概念的可视化表示形式。它们通常用于为控件和命令添加含义。实际或符号化的视觉对象使用户能够以与标记帮助用户浏览其环境的相同方式浏览 UI。这些视觉对象应简单明了，并且只包含所需的详细信息，以使客户能够快速分析他们在选择控件时将会发生的操作。

Office 功能区界面具有标准的视觉样式。如果你正在为 Office 功能区设计外接程序命令，请遵循我们的[图标准则](design-icons.md)。这可以确保一致性并熟悉各个 Office 应用程序。这些准则将有助于你为解决方案设计一组适合作为 Office 固有部分的 PNG 资产。

许多 HTML 容器包含带有插图的控件。使用 Office UI Fabric 的自定义字体在外接程序中呈现 Office 样式图标。Fabric 的图标字体包含很多针对可缩放的常见 Office 隐喻、颜色和样式的字形以满足你的需要。如果你有带自己图标集的现有视觉语言，则可在 HTML 画布中随意使用。构建自己带标准图标集的品牌的连续性是任何设计语言的重要组成部分。请注意避免与 Office 隐喻产生冲突导致客户混淆。

将下列一般原则应用于图标：

* 请勿在 Office 功能区或关联菜单中改变外接程序命令的 Office UI Fabric 用途。Fabric 图标风格不同，不能匹配。
* 使用 Office 图标语言来表示行为或概念。
* 将画笔等公用 Office 视觉隐喻重用于格式或用于查找的放大镜。
* 不得对不相关的操作误用隐喻。对不同的行为或概念使用相同的视觉效果可能会让用户感到困惑。


## <a name="see-also"></a>另请参阅

- [Office 加载项设计指南](add-in-design-guidelines.md)
- [在 Office 加载项中使用动作](using-motion-office-addins.md)
