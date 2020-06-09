---
title: 清单文件中的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612310"
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
