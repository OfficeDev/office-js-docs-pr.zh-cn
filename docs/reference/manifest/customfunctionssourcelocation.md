---
title: 清单文件中的 SourceLocation 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432404"
---
# <a name="sourcelocation-element"></a>SourceLocation 元素

定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。

## <a name="attributes"></a>属性

| **属性** | **必需** | **说明**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | 是          | 清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<SourceLocation resid="pageURL"/>
```