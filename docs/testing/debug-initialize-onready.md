---
title: 调试初始化和 onReady 方法
description: 了解如何调试 Office.initialize 和 Office.onReady 方法。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: ed6e69a52f3f4534db075daf62c171d4806e89d4
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797701"
---
# <a name="debug-the-initialize-and-onready-methods"></a>调试初始化和 onReady 方法

> [!NOTE]
> 本文假设你熟悉 [如何初始化 Office 加载项](../develop/initialize-add-in.md)。

调试 [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) 和 [Office.onReady](/javascript/api/office#office-office-onready-function(1)) 方法的悖论是调试器只能附加到正在运行的进程，但这些方法会在加载项运行时进程启动时立即运行，然后调试器才能附加。 在大多数情况下，在附加调试器后重启加载项没有帮助，因为重启外接程序会关闭原始运行时进程 *和附加的调试器* ，并启动一个不附加调试器的新进程。

幸运的是，存在异常。 可以使用Office web 版调试这些方法，并执行以下步骤。

1. 在Office web 版中旁加载并运行加载项。 这通常通过打开加载项的任务窗格或运行 [函数命令](../design/add-in-commands.md#types-of-add-in-commands)来完成。 *外接程序在整体浏览器进程中运行，而不是像桌面 Office 那样单独运行。*
1. 打开浏览器的开发人员工具。 这通常通过按 F12 来完成。 工具中的调试器附加到浏览器进程。
1. 根据需要将断点应用于或`Office.onReady`方法中的`Office.initialize`代码。
1. 像在步骤 1 中一样 *，重新启动加载项的任务窗格或函数命令*。 此操作 *不会* 关闭浏览器进程或调试器。 该 `Office.initialize` 或 `Office.onReady` 方法再次运行，处理在断点上停止。

> [!TIP]
> 有关更多详细信息，请参阅[Office web 版中的调试加载项](debug-add-ins-in-office-online.md)。 