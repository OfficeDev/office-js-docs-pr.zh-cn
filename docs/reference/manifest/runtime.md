---
title: 清单文件中的运行时
description: Runtime 元素将您的外接程序配置为对其功能区、任务窗格和自定义函数使用共享的 JavaScript 运行时。
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217758"
---
# <a name="runtime-element"></a>Runtime 元素

元素的子元素 [`<Runtimes>`](runtimes.md) 。 此元素将您的外接程序配置为使用共享的 JavaScript 运行时，以便功能区、任务窗格和自定义函数在同一运行时中运行。 有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

**外接程序类型：** 任务窗格

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
|  **生存时间 = "long"**  |  是  | 应始终是 `long` ，如果您想要为 Excel 加载项使用共享运行时。 |
|  **resid**  |  是  | 指定您的外接程序的 HTML 页面的 URL 位置。 `resid`必须与 `id` `Url` 元素中元素的属性相匹配 `Resources` 。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
