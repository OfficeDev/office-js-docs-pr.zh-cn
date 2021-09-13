---
title: 清单文件中的 Metadata 元素
description: Metadata 元素定义自定义函数在元数据Excel。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6f58b00bb13bde1e2b1742462716119b8b6d369d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152298"
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
