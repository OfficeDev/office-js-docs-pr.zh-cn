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
# <a name="sourcelocation-element"></a><span data-ttu-id="0c4da-102">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="0c4da-102">SourceLocation element</span></span>

<span data-ttu-id="0c4da-103">定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。</span><span class="sxs-lookup"><span data-stu-id="0c4da-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="0c4da-104">属性</span><span class="sxs-lookup"><span data-stu-id="0c4da-104">Attributes</span></span>

| <span data-ttu-id="0c4da-105">**属性**</span><span class="sxs-lookup"><span data-stu-id="0c4da-105">**Attribute**</span></span> | <span data-ttu-id="0c4da-106">**必需**</span><span class="sxs-lookup"><span data-stu-id="0c4da-106">**Required**</span></span> | <span data-ttu-id="0c4da-107">**说明**</span><span class="sxs-lookup"><span data-stu-id="0c4da-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="0c4da-108">resid</span><span class="sxs-lookup"><span data-stu-id="0c4da-108">resid</span></span>         | <span data-ttu-id="0c4da-109">是</span><span class="sxs-lookup"><span data-stu-id="0c4da-109">Yes</span></span>          | <span data-ttu-id="0c4da-110">清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。</span><span class="sxs-lookup"><span data-stu-id="0c4da-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="0c4da-111">子元素</span><span class="sxs-lookup"><span data-stu-id="0c4da-111">Child elements</span></span>

<span data-ttu-id="0c4da-112">无</span><span class="sxs-lookup"><span data-stu-id="0c4da-112">None</span></span>

## <a name="example"></a><span data-ttu-id="0c4da-113">示例</span><span class="sxs-lookup"><span data-stu-id="0c4da-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```