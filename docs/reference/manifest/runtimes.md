---
title: 清单文件中的运行时
description: 运行时元素指定外接程序的运行时。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: ef00bea317ae479d912b3a02f269ef97045b015d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608095"
---
# <a name="runtimes-element"></a>运行时元素

指定外接程序的运行时。 元素的子 [`<Host>`](host.md) 元素。

> [!NOTE]
> 在 Windows 上的 Office 中运行时，外接程序使用 Internet Explorer 11 浏览器。

在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。 有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

在 Outlook 中，此元素启用基于事件的加载项激活。 有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。

**外接类型：** 任务窗格、邮件

> [!IMPORTANT]
> **Excel**：共享运行时目前仅适用于 Windows 中的 Excel。
>
> **Outlook**：基于事件的激活功能当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅适用于 web 上的 Outlook。 有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>包含于

[Host](host.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  Description  |
|:-----|:-----|:-----|
| [运行时](runtime.md) | 是 |  外接程序的运行时。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtime.md)
