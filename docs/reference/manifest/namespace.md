---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在 Excel 中使用的命名空间。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612232"
---
# <a name="namespace-element"></a>Namespace 元素

定义 Excel 中的自定义函数使用的命名空间。

## <a name="attributes"></a>属性

|  属性  |  必需  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  否  | 应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<Namespace resid="namespace" />
```
