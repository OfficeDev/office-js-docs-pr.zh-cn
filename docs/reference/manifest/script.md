---
title: 清单文件中的 Script 元素
description: Script 元素定义自定义函数在自定义脚本Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51902864081e135faed778de1bc6fdee15d67490de8eabc9febf493cb0c09889
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095039"
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
