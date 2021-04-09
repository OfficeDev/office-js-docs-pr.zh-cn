---
title: 清单文件中运行时
description: Runtime 元素将外接程序配置为将共享 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652242"
---
# <a name="runtime-element"></a>运行时元素

将外接程序配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。 元素的 [`<Runtimes>`](runtimes.md) 子元素。

**外接程序类型：** 任务窗格、邮件

[!include[Runtimes support](../../includes/runtimes-note.md)]

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
|  **resid**  |  是  | 指定外接程序的 HTML 页面的 URL 位置。 `resid`不能超过 32 个字符，并且必须与 元素中的 `id` `Url` 元素的 属性 `Resources` 匹配。 |
|  **lifetime**  |  否  | 的默认值是 `lifetime` `short` ，不需要指定。 Outlook 外接程序仅使用 `short` 值。 如果要在 Excel 加载项中使用共享运行时，请显式将值设置为 `long` 。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [配置 Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)
