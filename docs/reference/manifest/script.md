---
title: 清单文件中的 Script 元素
description: Script 元素定义自定义函数在自定义脚本Excel。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 259976f752cf3fca72c5012bedd92b9bf021f6aa
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990668"
---
# <a name="script-element"></a>Script 元素

定义 Excel 中的自定义函数所使用的脚本设置。

**外接程序类型：** 自定义函数

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
