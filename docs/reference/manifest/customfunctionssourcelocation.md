---
title: 清单文件中自定义函数的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: b18a340d4dd4403b1e5fd2c7d8868a820eef5a241ac3d666926d8f2cb49fcc09
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098297"
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
