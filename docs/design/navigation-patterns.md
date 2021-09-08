---
title: Office 外接程序的导航模式
description: 了解使用命令栏、选项卡栏和后退按钮的最佳实践，以设计加载项Office导航。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938931"
---
# <a name="navigation-patterns"></a>导航模式

可以通过特定命令类型和指定的屏幕区域访问外接程序的主要功能。 导航直观明了，可提供上下文并允许用户在外接程序中轻松移动，这些非常重要。

## <a name="best-practices"></a>最佳做法

| 允许事项    | 禁止事项 |
| :---- | :---- |
| 确保为用户提供清晰的可视化导航选项。 | 不要使用非标准 UI，使导航过程变得复杂。
| 使用以下组件（如适用）允许用户在加载程序中导航。 | 不要让用户难以知悉其当前在外接程序中所处的位置或上下文

## <a name="command-bar"></a>命令栏

CommandBar 是任务窗格中的一个图面，其中包含对它所在的窗口、面板或父区域的内容进行操作的命令。 可选功能包括汉堡菜单访问点、搜索和侧命令。

![插图显示桌面应用程序任务Office内的命令栏。 此示例显示紧接在外接程序名称下方的命令栏，其中包括汉堡包菜单和搜索。](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>选项卡栏

选项卡栏显示使用具有垂直堆叠文本和图标的按钮的导航。 使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。

![插图显示桌面应用程序任务Office内的选项卡栏。 此示例显示紧接在外接程序名称下方的选项卡栏，其选项卡具有"Home"、"设置"、"Favorites"和"Account"选项卡。](../images/add-in-tab-bar.png)

## <a name="back-button"></a>“返回”按钮

"后退"按钮允许用户从向下钻取导航操作中恢复。 此模式有助于确保用户遵循一系列有序的步骤。

![插图显示桌面应用程序任务Office内的后退按钮。 本示例在加载项名称的下方左上方显示一个后退按钮。](../images/add-in-back-button.png)
