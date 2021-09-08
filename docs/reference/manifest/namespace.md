---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在自定义函数中Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937694"
---
# <a name="namespace-element"></a>Namespace 元素

定义 Excel 中的自定义函数使用的命名空间。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  否  | 应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。 不能超过 32 个字符。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<Namespace resid="namespace" />
```
