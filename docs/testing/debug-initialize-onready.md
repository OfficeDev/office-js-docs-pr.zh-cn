---
title: 调试 initialize 和 onReady 函数
description: 了解如何调试 Office.initialize 和 Office.onReady 函数。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dca551d8a016e7aad16cfdc02590f0a51455852
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423249"
---
# <a name="debug-the-initialize-and-onready-functions"></a>调试 initialize 和 onReady 函数

> [!NOTE]
> 本文假设你熟悉 [如何初始化 Office 加载项](../develop/initialize-add-in.md)。

调试 [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) 和 [Office.onReady 函](/javascript/api/office#office-office-onready-function(1)) 数的悖论是调试器只能附加到正在运行的进程，但这些函数在加载项的运行时进程启动时立即运行，然后调试程序才能附加。 在大多数情况下，在附加调试器后重启加载项没有帮助，因为重启外接程序会关闭原始运行时进程 *和附加的调试器* ，并启动一个不附加调试器的新进程。

幸运的是，存在异常。 可以使用Office web 版调试这些函数，并执行以下步骤。

1. 在Office web 版中旁加载并运行加载项。 这通常通过打开加载项的任务窗格或运行 [函数命令](../design/add-in-commands.md#types-of-add-in-commands)来完成。 *外接程序在整体浏览器进程中运行，而不是像桌面 Office 那样单独运行。*
1. 打开浏览器的开发人员工具。 这通常通过按 F12 来完成。 工具中的调试器附加到浏览器进程。
1. 根据需要将断点应用于或`Office.onReady`函数中的`Office.initialize`代码。
1. 像在步骤 1 中一样 *，重新启动加载项的任务窗格或函数命令*。 此操作 *不会* 关闭浏览器进程或调试器。 该 `Office.initialize` 或 `Office.onReady` 函数再次运行，处理在断点上停止。

> [!TIP]
> 有关更多详细信息，请参阅[Office web 版中的调试加载项](debug-add-ins-in-office-online.md)。

## <a name="see-also"></a>另请参阅

- [Office 加载项中的运行时](runtimes.md)
