---
title: 清单文件中自定义函数的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937451"
---
# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 元素 (自定义函数) 

定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。

## <a name="attributes"></a>属性

| 属性 | 必需 | 说明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | 是      | 清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。 不能超过 32 个字符。 |

## <a name="child-elements"></a>子元素

无

## <a name="example"></a>示例

```xml
<SourceLocation resid="pageURL"/>
```
