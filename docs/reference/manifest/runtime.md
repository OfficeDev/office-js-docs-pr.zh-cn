---
title: 清单文件中的运行时
description: Runtime 元素将您的外接程序配置为对其各个组件使用共享的 JavaScript 运行时，例如，功能区、任务窗格、自定义函数。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159365"
---
# <a name="runtime-element-preview"></a>Runtime 元素（预览）

将您的外接程序配置为使用共享的 JavaScript 运行时，以便在同一运行时中运行各种组件。 元素的子 [`<Runtimes>`](runtimes.md) 元素。

在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。 有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

在 Outlook 中，此元素启用基于事件的加载项激活。 有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。

**外接类型：** 任务窗格、邮件

> [!IMPORTANT]
> **Outlook**：基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅适用于 web 上的 outlook。 有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

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
|  **resid**  |  是  | 指定您的外接程序的 HTML 页面的 URL 位置。 `resid`必须与 `id` `Url` 元素中元素的属性相匹配 `Resources` 。 |
|  **lifetime**  |  否  | 的默认值 `lifetime` 是 `short` ，不需要指定。 Outlook 外接程序仅使用 `short` 值。 如果要在 Excel 外接程序中使用共享运行时，请将值显式设置为 `long` 。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
