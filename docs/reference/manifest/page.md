---
title: 清单文件中的 Page 元素
description: Page 元素定义自定义函数在自定义页面中使用的 HTML Excel。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6bde3ba86270874b1d9059b2f1c44952241bf00f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152646"
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
