---
title: 清单文件中的 SourceLocation 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450686"
---
# <a name="sourcelocation-element"></a>SourceLocation 元素

定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。

## <a name="attributes"></a>属性

| **属性** | **必需** | **描述**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | 是          | 清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<SourceLocation resid="pageURL"/>
```
