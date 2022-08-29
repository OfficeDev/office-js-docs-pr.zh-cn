---
title: 使用非共享运行时调试函数命令
description: 了解如何调试函数命令。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: d2be148c05f88837610b8563c2e61618d1c37775
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423200"
---
# <a name="debug-a-function-command-with-a-non-shared-runtime"></a>使用非共享运行时调试函数命令

> [!IMPORTANT]
> 如果外接程序 [配置为使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)，则调试函数命令后面的代码，就像任务窗格后面的代码一样。 请参阅 [调试 Office 加载项](debug-add-ins-overview.md) ，并注意加载项中具有 [共享运行时的](runtimes.md#shared-runtime) 函数命令 *不是* 本文所述的特殊情况。 

> [!NOTE]
> 本文假定你熟悉 [函数命令](../design/add-in-commands.md#types-of-add-in-commands)。

函数命令没有 UI，因此无法将调试器附加到在桌面 Office 上运行函数的过程。  (在 Windows 上开发的 Outlook 加载项是一个例外。 请参阅本文稍后[在 Windows 上的 Outlook 外接程序中的调试](#debug-function-commands-in-outlook-add-ins-on-windows)函数命令。) 在具有非共享运行时的外接程序中，必须调试函数在整个浏览器进程中运行的Office web 版。 使用以下步骤。

1. 在Office web 版中旁加载外接程序，然后选择运行函数命令的按钮或菜单项。 这是加载函数命令的代码文件所必需的。 
1. 打开浏览器的开发人员工具。 这通常通过按 F12 来完成。 工具中的调试器附加到浏览器进程。
1. 根据函数命令的需要将断点应用于代码。
1. 重新运行函数命令。 进程在断点上停止。 

> [!TIP]
> 有关更多详细信息，请参阅[Office web 版中的调试加载项](debug-add-ins-in-office-online.md)。

## <a name="debug-function-commands-in-outlook-add-ins-on-windows"></a>在 Windows 上的 Outlook 加载项中调试函数命令

如果开发计算机是 Windows，可以通过一种方法在 Outlook 桌面上调试函数命令。 请参阅 [Outlook 加载项中的调试函数命令](../outlook/debug-ui-less.md)。

## <a name="see-also"></a>另请参阅

- [Office 加载项中的运行时](runtimes.md)
