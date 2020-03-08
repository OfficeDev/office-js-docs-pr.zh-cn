---
title: 清单文件中的运行时（预览）
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561826"
---
# <a name="runtime-element-preview"></a>Runtime 元素（预览）

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

[`<Runtimes>`](runtimes.md)元素的子元素。 此元素将您的外接程序配置为使用共享的 JavaScript 运行时，以便功能区、任务窗格和自定义函数在同一运行时中运行。 有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

**外接程序类型：** 任务窗格

> [!IMPORTANT]
> 共享运行时当前处于预览阶段，仅适用于 Windows 上的 Excel。 若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。

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
|  **生存时间 = "long"**  |  是  | 应始终是`long` ，如果您想要为 Excel 加载项使用共享运行时。 |
|  **resid**  |  是  | 指定您的外接程序的 HTML 页面的 URL 位置。 `resid`必须与`Resources`元素中`id` `Url`元素的属性相匹配。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
