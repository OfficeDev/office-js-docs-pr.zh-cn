---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917084"
---
# <a name="runtimes-element"></a>Runtimes 元素

指定外接程序的运行时。 元素的 [`<Host>`](host.md) 子元素。

> [!NOTE]
> When running in Office on Windows， an add-in that has a element in its manifest does notnecessarily `<Runtimes>` run in the same webview control as it otherwise would. 有关 Windows 和 Office 版本如何确定正常使用的 Webview 控件的信息，请参阅 [Office 外接程序使用的浏览器](../../concepts/browsers-used-by-office-web-add-ins.md)。如果满足针对将 Microsoft Edge 与 WebView2 一 (基于 Chromium) 的条件，则无论外接程序是否具有 元素，外接程序都使用该 `<Runtimes>` 浏览器。 但是，当不满足这些条件时，具有 元素的外接程序始终使用 `<Runtimes>` Internet Explorer 11，无论 Windows 或 Microsoft 365 版本如何。

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
| [运行时](runtime.md) | 是 |  加载项的运行时。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtime.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [配置 Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)
