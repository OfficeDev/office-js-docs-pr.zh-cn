---
title: 清单文件中的 Namespace 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432324"
---
# <a name="namespace-element"></a>Namespace 元素

定义 Excel 中的自定义函数使用的命名空间。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  是  | 应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<Namespace resid="namespace" />
```
