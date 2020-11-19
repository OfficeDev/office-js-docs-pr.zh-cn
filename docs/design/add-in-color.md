---
title: Office 外接程序的颜色准则
description: 了解如何使用 Office 外接程序的 UI 中的颜色。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 99eef66ec5ed1cb421d4d8cef7e20d8b19a0ee3d
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132184"
---
# <a name="color"></a>颜色

颜色通常用于强调品牌和强化可视化层次结构。 它可以帮助标识接口，也可以指导客户完成体验。 在 Office 中，颜色用于相同目标，但应有目的地最小限度地应用它。 它不会过度使用客户内容。 即使每个 Office 应用程序被标上了自己的主色，但还是很少用到。

![显示 Office、Excel、Word 和 PowerPoint 的配色方案的图示。 Office 的主要颜色为黑色和白色，次要颜色为浅灰色、深灰色和橙色。 Excel 的主颜色为绿色，Word 为蓝色，而 PowerPoint 为橙色。](../images/office-addins-color-schemes.png)

Office UI Fabric 包含一组默认主题颜色。当 Fabric 作为组件应用于 Office 外接程序或应用于布局时，相同的目标适用。颜色应传达层次结构，有目的地指导客户操作而不会干扰内容。Fabric 主题颜色可以向整体界面引入新的个性色。此新的个性色可能会与 Office 应用程序品牌产生冲突并干扰层次结构。换句话说，Fabric 在外接程序内部使用时可能会向整体界面引入新的个性色。此新的个性色可能会分散用户注意力并干扰整个层次结构。寻找避免冲突和干扰的方法。使用中性个性色或覆盖 Fabric 主题颜色，以匹配 Office 应用程序品牌或你自己的品牌颜色。

Office 应用程序使客户能够通过应用 Office UI 主题个性化设置其界面。 客户可以在四个 UI 主题中进行选择来改变背景样式以及 Word、PowerPoint、Excel 和 Office 套件中其他应用程序的按钮。 若要使您的加载项感觉像是 Office 的自然部件，并对个性化做出响应，请使用我们的主题 Api。 例如，任务窗格背景颜色在某些主题中切换到深灰色。 我们的主题 API 允许你照做并调整前景文本，以确保[辅助功能](../design/accessibility-guidelines.md)。

> [!NOTE]
> - 对于邮件和任务窗格外接程序，请使用 [Context.officeTheme](/javascript/api/office/office.context) 元素匹配 Office 应用程序的主题。 此 API 目前在 Office 2016 或更高版本中可用。
> - 对于 PowerPoint 内容加载项，请参阅[在 PowerPoint 加载项中使用 Office 主题](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)。

将下列一般原则应用于颜色：

- 尽量少使用颜色来显示层次结构和强调品牌。
- 过度使用单个应用于交互式和非交互式元素的个性色可能会导致混乱。例如，避免将相同颜色用于导航菜单中的选定和未选定项。
- 避免与 Office 品牌应用颜色产生不必要的冲突。
- 使用自己的品牌颜色来生成与服务或公司的关联。
- 确保可以访问所有文本。 请确保前景文本和背景的速率为4.5：1对比度。
- 注意色盲群体。 不要仅使用颜色指示交互性和层次结构。
- 请参阅 [图标指南](../design/add-in-icons.md) ，了解有关使用 Office 图标颜色调色板设计外接命令图标的详细信息。
