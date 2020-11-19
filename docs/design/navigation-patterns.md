---
title: Office 外接程序的导航模式
description: 了解使用命令栏、选项卡栏和后退按钮的最佳实践，以设计 Office 外接程序的导航。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132030"
---
# <a name="navigation-patterns"></a>导航模式

可以通过特定命令类型和指定的屏幕区域访问外接程序的主要功能。 导航直观明了，可提供上下文并允许用户在外接程序中轻松移动，这些非常重要。

## <a name="best-practices"></a>最佳做法

| 允许事项    | 禁止事项 |
| :---- | :---- |
| 确保为用户提供清晰的可视化导航选项。 | 不要使用非标准 UI，使导航过程变得复杂。
| 使用以下组件（如适用）允许用户在加载程序中导航。 | 不要让用户难以知悉其当前在外接程序中所处的位置或上下文

## <a name="command-bar"></a>命令栏

命令栏是任务窗格中的一个图面，其中驻留了在其驻留的窗口、面板或父区域的内容上运行的命令。 可选功能包括汉堡菜单访问点、搜索和侧命令。

![图示显示在 "Office 桌面应用程序" 任务窗格中的命令栏。 本示例将一个命令栏显示在包含汉堡菜单和搜索的外接程序名称的正下方。](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>选项卡栏

选项卡栏显示了使用垂直堆叠文本和图标的按钮的导航。 使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。

![图示显示在 "Office 桌面应用程序" 任务窗格中的选项卡栏。 本示例在外接姓名下方显示一个选项卡栏，其中包含 "主页"、"设置"、"收藏夹" 和 "帐户" 选项卡。](../images/add-in-tab-bar.png)

## <a name="back-button"></a>“返回”按钮

"后退" 按钮允许用户从深化导航操作中恢复。 此模式有助于确保用户遵循一系列有序的步骤。

![显示 Office 桌面应用程序任务窗格中的 "后退" 按钮的图示。 本示例显示一个 "后退" 按钮，该按钮位于加载项名称的左上角。](../images/add-in-back-button.png)
