---
title: 清单文件中运行时
description: Runtimes 元素指定加载项的运行时。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789161"
---
# <a name="runtimes-element"></a>Runtimes 元素

指定加载项的运行时。 元素的 [`<Host>`](host.md) 子元素。

> [!NOTE]
> When running in Office on Windows， your add-in uses the Internet Explorer 11 browser.

在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。 有关详细信息，请参阅配置 [Excel 加载项以使用共享的 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

在 Outlook 中，此元素启用基于事件的外接程序激活。 有关详细信息，请参阅配置 [Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)。

**加载项类型：** 任务窗格、邮件

> [!IMPORTANT]
> **Outlook：** 基于事件的激活 [功能目前处于](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 预览阶段，仅在 Outlook 网页版中可用。 有关详细信息，请参阅 [如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

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
