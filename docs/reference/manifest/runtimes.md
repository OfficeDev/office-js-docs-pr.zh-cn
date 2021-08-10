---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: e5ec70449d3984671d507131ac8d4fc0f7617cdcda1ad8f99b4f4bf52773aded
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091624"
---
# <a name="runtimes-element"></a>Runtimes 元素

指定外接程序的运行时。 元素的 [`<Host>`](host.md) 子元素。

> [!NOTE]
> 当在 Office Windows 中运行时，其清单中具有 元素的加载项不必像否则一样在同一 `<Runtimes>` Webview 控件中运行。 有关 web 视图和 Windows Office版本如何确定通常使用的 Webview 控件Office[请参阅](../../concepts/browsers-used-by-office-web-add-ins.md)浏览器。如果满足将 Microsoft Edge与 WebView2 (Chromium) 的条件，则无论外接程序是否具有 元素，外接程序都使用该 `<Runtimes>` 浏览器。 但是，当不满足这些条件时，具有 元素的外接程序始终使用 Internet Explorer 11，而不考虑 Windows `<Runtimes>` 或 Microsoft 365 版本。

**外接程序类型：** 任务窗格、邮件

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>包含于

[Host](host.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [运行时](runtime.md) | 是 |  加载项的运行时。 **重要** 提示：目前，只能定义一 `<Runtime>` 个元素。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtime.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md)
