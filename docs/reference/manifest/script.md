---
title: 清单文件中的 Script 元素
description: Script 元素定义自定义函数在 Excel 中使用的脚本设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f05fc85bd0454c340f4352bb73f299b9e7730224
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720412"
---
# <a name="script-element"></a>Script 元素

定义 Excel 中的自定义函数所使用的脚本设置。

## <a name="attributes"></a>属性

无

## <a name="child-elements"></a>子元素

|元素  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 包含自定义函数所使用的 JavaScript 文件的资源 ID 的字符串。|

## <a name="example"></a>示例

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
