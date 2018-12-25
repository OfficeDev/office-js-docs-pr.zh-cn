---
title: 清单文件中的 Page 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433731"
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
