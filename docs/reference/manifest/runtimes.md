---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 758bb7b830009d6691190a0279440a52da724624
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138602"
---
# <a name="runtimes-element"></a>Runtimes 元素

指定外接程序的运行时。 元素的 [`<Host>`](host.md) 子元素。

> [!NOTE]
> 当在 Office Windows 中运行时，其清单中具有元素的外接程序不必像否则一样在同一 `<Runtimes>` Webview 控件中运行。 有关网站和加载项Windows Office哪些 Webview 控件是正常使用的信息，请参阅浏览器Office[加载项。](../../concepts/browsers-used-by-office-web-add-ins.md)如果满足针对将 Microsoft Edge 与 WebView2 (Chromium的条件) ，则无论外接程序是否具有 元素，外接程序都使用该 `<Runtimes>` 浏览器。 但是，当不满足这些条件时，具有 元素的外接程序始终使用 Internet Explorer 11，而不考虑 Windows `<Runtimes>` 或 Microsoft 365 版本。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

 - 任务窗格 1.0
 - 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (仅在任务窗格外接程序中使用时) 

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
