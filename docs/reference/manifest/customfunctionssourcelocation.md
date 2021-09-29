---
title: 清单文件中自定义函数的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5f2d881f31f4e46e7f5bb8ab30d78abd0e9b7200
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990682"
---
# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 元素 (自定义函数) 

定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。

**外接程序类型：** 自定义函数

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
