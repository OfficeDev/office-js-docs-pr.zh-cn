---
title: 清单文件中的运行时
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111168"
---
# <a name="runtime-element"></a>Runtime 元素

此功能处于预览阶段。 [`<Runtimes>`](runtime.md)元素的子元素。 此元素有助于在 Excel 自定义函数和外接程序的任务窗格之间共享全局数据和函数调用。 

## <a name="contained-in"></a>包含于

-[时](runtimes.md)

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **生存时间 = "long"**  |  是  | 如果希望 Excel 自定义函数在外接程序的任务窗格关闭时正常工作，应始终将其列为长。 |
|  **resid**  |  是  | 如果用于 Excel 自定义函数，则`resid`应指向`TaskPaneAndCustomFunction.Url`。 |

## <a name="see-also"></a>另请参阅

-[语言](runtime.md)
