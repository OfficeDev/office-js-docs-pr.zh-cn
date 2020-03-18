---
title: 清单文件中的 Page 元素
description: Page 元素定义了自定义函数在 Excel 中使用的 HTML 页面设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720482"
---
# <a name="page-element"></a>Page 元素

定义 Excel 中的自定义函数所使用的 HTML 页面设置。

## <a name="attributes"></a>属性

无

## <a name="child-elements"></a>子元素

|  元素  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。 |

## <a name="example"></a>示例

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
