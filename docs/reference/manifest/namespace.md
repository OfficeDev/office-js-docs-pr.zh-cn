---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在自定义函数中Excel。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 3a5afed3d55bde7e9735df534215f96ae1ba7bd3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152664"
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
