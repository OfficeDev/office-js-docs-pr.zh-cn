---
title: Office 外接程序的导航模式
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: f0482f7742c6fab97fe563d61d21135c072f8f8f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449131"
---
# <a name="navigation-patterns"></a>导航模式

可以通过特定命令类型和指定的屏幕区域访问外接程序的主要功能。 导航直观明了，可提供上下文并允许用户在外接程序中轻松移动，这些非常重要。

## <a name="best-practices"></a>最佳做法

| 允许事项    | 禁止事项 |
| :---- | :---- |
| 确保为用户提供清晰的可视化导航选项。 | 不要使用非标准 UI，使导航过程变得复杂。
| 使用以下组件（如适用）允许用户在加载程序中导航。 | 不要让用户难以知悉其当前在外接程序中所处的位置或上下文



## <a name="command-bar"></a>命令栏

命令栏是一个图面，其中包含在其驻留的窗口、面板或父区域内容上运行的命令。 可选功能包括汉堡菜单访问点、搜索和侧命令。

![命令 - 桌面任务窗格规范](../images/add-in-command-bar.png)



## <a name="tab-bar"></a>选项卡栏

显示使用具有垂直堆叠文本和图标的按钮进行导航。 使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。

![选项卡栏 - 桌面任务窗格规范](../images/add-in-tab-bar.png)


## <a name="back-button"></a>“返回”按钮

“返回”按钮使用户能够恢复向下钻取导航操作。 此模式有助于确保用户遵循一系列有序的步骤。  

![“返回”按钮 - 桌面任务窗格规范](../images/add-in-back-button.png)
