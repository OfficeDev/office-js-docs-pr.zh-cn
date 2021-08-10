---
title: 清单文件中的 Metadata 元素
description: Metadata 元素定义自定义函数在自定义函数中Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d6b7af8988265baf8fbdea51e1414646f88868ede76ed7194c47db1eb874608d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092485"
---
# <a name="metadata-element"></a>Metadata 元素

定义 Excel 中的自定义函数所使用的元数据设置。

## <a name="attributes"></a>属性

无

## <a name="child-elements"></a>子元素

|  元素  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。 |

## <a name="example"></a>示例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
