---
title: 从任务窗格附加调试器
description: 了解如何从任务窗格附加调试器。
ms.date: 01/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0363b7966ab3da11167cb4b0cd324df28fd9efb3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744750"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>从任务窗格附加调试器

在某些环境中，调试器可以附加到Office运行的加载项上。 当您要调试已位于暂存或生产中的外接程序时，这非常有用。 如果您仍在开发和测试外接程序，请参阅调试外接程序Office[概述](debug-add-ins-overview.md)。

本文中所述的技术只能在满足以下条件时使用。

- 加载项在加载项的 Office 中Windows。
- 计算机使用使用基于 Edge Windows Office Webview 控件 (Chromium 的) 版本和) 版本。 若要确定你使用的浏览器，请参阅浏览器[Office外接程序](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

若要启动调试器，请选择任务窗格的右上角以激活"个性"菜单 (如下图所示的红色圆圈) 。

!["附加调试器"菜单的屏幕截图。](../images/attach-debugger.png)

选择“**附加调试器**”。 这将启动基于Microsoft Edge (Chromium的) 开发人员工具。 使用使用基于 web 的应用程序中的开发人员工具调试Microsoft Edge (Chromium[中所述) ](debug-add-ins-using-devtools-edge-chromium.md)。

## <a name="see-also"></a>另请参阅

- [调试 Office 加载项概述](debug-add-ins-overview.md)
