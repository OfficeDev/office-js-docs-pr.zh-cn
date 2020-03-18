---
title: 清单文件中的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720685"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="22ecf-103">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="22ecf-103">SourceLocation element</span></span>

<span data-ttu-id="22ecf-104">定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。</span><span class="sxs-lookup"><span data-stu-id="22ecf-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="22ecf-105">属性</span><span class="sxs-lookup"><span data-stu-id="22ecf-105">Attributes</span></span>

| <span data-ttu-id="22ecf-106">**属性**</span><span class="sxs-lookup"><span data-stu-id="22ecf-106">**Attribute**</span></span> | <span data-ttu-id="22ecf-107">**必需**</span><span class="sxs-lookup"><span data-stu-id="22ecf-107">**Required**</span></span> | <span data-ttu-id="22ecf-108">**描述**</span><span class="sxs-lookup"><span data-stu-id="22ecf-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="22ecf-109">resid</span><span class="sxs-lookup"><span data-stu-id="22ecf-109">resid</span></span>         | <span data-ttu-id="22ecf-110">是</span><span class="sxs-lookup"><span data-stu-id="22ecf-110">Yes</span></span>          | <span data-ttu-id="22ecf-111">清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。</span><span class="sxs-lookup"><span data-stu-id="22ecf-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="22ecf-112">子元素</span><span class="sxs-lookup"><span data-stu-id="22ecf-112">Child elements</span></span>

<span data-ttu-id="22ecf-113">无</span><span class="sxs-lookup"><span data-stu-id="22ecf-113">None</span></span>

## <a name="example"></a><span data-ttu-id="22ecf-114">示例</span><span class="sxs-lookup"><span data-stu-id="22ecf-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
