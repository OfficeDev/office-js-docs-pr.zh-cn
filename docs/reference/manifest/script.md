---
title: 清单文件中的 Script 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433136"
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
