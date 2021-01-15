---
title: 清单文件中自定义函数的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771380"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="b4804-103">SourceLocation 元素 (自定义函数) </span><span class="sxs-lookup"><span data-stu-id="b4804-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="b4804-104">定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。</span><span class="sxs-lookup"><span data-stu-id="b4804-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="b4804-105">属性</span><span class="sxs-lookup"><span data-stu-id="b4804-105">Attributes</span></span>

| <span data-ttu-id="b4804-106">属性</span><span class="sxs-lookup"><span data-stu-id="b4804-106">Attribute</span></span> | <span data-ttu-id="b4804-107">必需</span><span class="sxs-lookup"><span data-stu-id="b4804-107">Required</span></span> | <span data-ttu-id="b4804-108">说明</span><span class="sxs-lookup"><span data-stu-id="b4804-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="b4804-109">resid</span><span class="sxs-lookup"><span data-stu-id="b4804-109">resid</span></span>     | <span data-ttu-id="b4804-110">是</span><span class="sxs-lookup"><span data-stu-id="b4804-110">Yes</span></span>      | <span data-ttu-id="b4804-111">清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。</span><span class="sxs-lookup"><span data-stu-id="b4804-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> <span data-ttu-id="b4804-112">不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4804-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="b4804-113">子元素</span><span class="sxs-lookup"><span data-stu-id="b4804-113">Child elements</span></span>

<span data-ttu-id="b4804-114">无</span><span class="sxs-lookup"><span data-stu-id="b4804-114">None</span></span>

## <a name="example"></a><span data-ttu-id="b4804-115">示例</span><span class="sxs-lookup"><span data-stu-id="b4804-115">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
