---
title: Office 外接程序的颜色准则
description: 了解如何在加载项的 UI Office颜色。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: b43d3144f24f7b90878bcabe12db492f8dbe4f6f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742805"
---
# <a name="color-guidelines-for-office-add-ins"></a>Office 外接程序的颜色准则

颜色通常用于强调品牌和强化可视化层次结构。 它可以帮助标识接口，也可以指导客户完成体验。 在 Office 中，颜色用于相同目标，但应有目的地最小限度地应用它。 它不会过度使用客户内容。 即使每个 Office 应用程序被标上了自己的主色，但还是很少用到。

![此图显示 Office、Excel、Word 和 PowerPoint 的配色方案。 主要颜色Office为黑白，次要颜色为浅灰色、深灰色和橙色。 文本的基准颜色Excel绿色，Word 为蓝色，PowerPoint橙色。](../images/office-addins-color-schemes.png)

[Fabric Core](fabric-core.md) 包括一组默认主题颜色。 在组件或布局中将 Fabric Core Office外接程序时，相同的目标适用。 颜色应传达层次结构，有目的地指导客户操作而不会干扰内容。 Fabric Core 主题颜色可以在整体界面中引入新的主题色。 此新的个性色可能会与 Office 应用程序品牌产生冲突并干扰层次结构。 换句话说，在加载项内使用时，Fabric Core 可以将新的强调文字颜色引入整个界面。 此新的个性色可能会分散用户注意力并干扰整个层次结构。 寻找避免冲突和干扰的方法。 使用中性主题或覆盖 Fabric Core 主题颜色来匹配Office 应用品牌或你自己的品牌颜色。

Office 应用程序使客户能够通过应用 Office UI 主题个性化设置其界面。 客户可以在四个 UI 主题中进行选择来改变背景样式以及 Word、PowerPoint、Excel 和 Office 套件中其他应用程序的按钮。 若要使加载项感觉自己就像是应用和Office的一部分，请使用我们"Theming"API。 例如，任务窗格背景颜色在某些主题中切换到深灰色。 我们的主题 API 允许你照做并调整前景文本，以确保[辅助功能](../design/accessibility-guidelines.md)。

> [!NOTE]
>
> - 对于邮件和任务窗格外接程序，请使用 [Context.officeTheme](/javascript/api/office/office.context) 元素匹配 Office 应用程序的主题。 此 API 当前在 Office 2016 或更高版本中可用。
> - 对于 PowerPoint 内容外接程序，请参阅[在 PowerPoint 外接程序中使用 Office 主题](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)。

应用以下颜色一般准则。

- 尽量少使用颜色来显示层次结构和强调品牌。
- 过度使用单个应用于交互式和非交互式元素的个性色可能会导致混乱。例如，避免将相同颜色用于导航菜单中的选定和未选定项。
- 避免与 Office 品牌应用颜色产生不必要的冲突。
- 使用自己的品牌颜色来生成与服务或公司的关联。
- 确保可以访问所有文本。 确保前景文本和背景之间的对比率为 4.5：1。
- 注意色盲群体。不要仅使用颜色指示交互性和层次结构。
- 请参阅[图标指南](../design/add-in-icons.md)，了解有关使用图标颜色调色板设计Office命令图标。
