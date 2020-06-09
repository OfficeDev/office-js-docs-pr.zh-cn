---
title: 清单文件中的 Metadata 元素
description: Metadata 元素定义自定义函数在 Excel 中使用的元数据设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611762"
---
# <a name="metadata-element"></a>Metadata 元素

定义 Excel 中的自定义函数所使用的元数据设置。

## <a name="attributes"></a>属性

无

## <a name="child-elements"></a>子元素

|  元素  |  必需  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。 |

## <a name="example"></a>示例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
