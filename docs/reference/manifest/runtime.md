---
title: 清单文件中运行时
description: 运行时元素将加载项配置为将共享的 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789182"
---
# <a name="runtime-element-preview"></a>运行时元素 (预览) 

将加载项配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。 元素的 [`<Runtimes>`](runtimes.md) 子元素。

在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。 有关详细信息，请参阅配置 [Excel 加载项以使用共享的 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

在 Outlook 中，此元素启用基于事件的外接程序激活。 有关详细信息，请参阅配置 [Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)。

**加载项类型：** 任务窗格、邮件

> [!IMPORTANT]
> **Outlook：** 基于事件的激活目前处于 [预览阶段，](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 仅在 Outlook 网页版中可用。 有关详细信息，请参阅 [如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>包含于

- [运行时](runtimes.md)

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **resid**  |  是  | 指定外接程序的 HTML 页面的 URL 位置。 `resid`不能超过 32 个字符，并且必须与元素 `id` `Url` 中的元素属性 `Resources` 匹配。 |
|  **lifetime**  |  否  | 默认值是 `lifetime` `short` ，不需要指定。 Outlook 外接程序仅使用 `short` 该值。 如果要在 Excel 加载项中使用共享运行时，请显式将值设置为 `long` 。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
