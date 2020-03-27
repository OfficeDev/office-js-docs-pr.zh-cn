---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在 Excel 中使用的命名空间。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978667"
---
# <a name="namespace-element"></a>Namespace 元素

定义 Excel 中的自定义函数使用的命名空间。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  否  | 应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<Namespace resid="namespace" />
```
