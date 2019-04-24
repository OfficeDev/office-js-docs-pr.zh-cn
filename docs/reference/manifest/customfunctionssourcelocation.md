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
# <a name="sourcelocation-element"></a><span data-ttu-id="74599-102">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="74599-102">SourceLocation element</span></span>

<span data-ttu-id="74599-103">定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。</span><span class="sxs-lookup"><span data-stu-id="74599-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="74599-104">属性</span><span class="sxs-lookup"><span data-stu-id="74599-104">Attributes</span></span>

| <span data-ttu-id="74599-105">**属性**</span><span class="sxs-lookup"><span data-stu-id="74599-105">**Attribute**</span></span> | <span data-ttu-id="74599-106">**必需**</span><span class="sxs-lookup"><span data-stu-id="74599-106">**Required**</span></span> | <span data-ttu-id="74599-107">**描述**</span><span class="sxs-lookup"><span data-stu-id="74599-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="74599-108">resid</span><span class="sxs-lookup"><span data-stu-id="74599-108">resid</span></span>         | <span data-ttu-id="74599-109">是</span><span class="sxs-lookup"><span data-stu-id="74599-109">Yes</span></span>          | <span data-ttu-id="74599-110">清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。</span><span class="sxs-lookup"><span data-stu-id="74599-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="74599-111">子元素</span><span class="sxs-lookup"><span data-stu-id="74599-111">Child elements</span></span>

<span data-ttu-id="74599-112">无</span><span class="sxs-lookup"><span data-stu-id="74599-112">None</span></span>

## <a name="example"></a><span data-ttu-id="74599-113">示例</span><span class="sxs-lookup"><span data-stu-id="74599-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
