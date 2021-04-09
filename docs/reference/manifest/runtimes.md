---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652230"
---
# <a name="runtimes-element"></a>Runtimes 元素

指定外接程序的运行时。 元素的 [`<Host>`](host.md) 子元素。

> [!NOTE]
> When running in Office on Windows， your add-in uses the Internet Explorer 11 browser.

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
