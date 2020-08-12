---
title: 清单文件中的自定义函数的 SourceLocation 元素
description: 定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641380"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="4ff8b-103">SourceLocation 元素 (自定义函数) </span><span class="sxs-lookup"><span data-stu-id="4ff8b-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="4ff8b-104">定义 Excel 中自定义函数所使用的 Script 或 Page 元素所需的资源的位置。</span><span class="sxs-lookup"><span data-stu-id="4ff8b-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="4ff8b-105">属性</span><span class="sxs-lookup"><span data-stu-id="4ff8b-105">Attributes</span></span>

| <span data-ttu-id="4ff8b-106">属性</span><span class="sxs-lookup"><span data-stu-id="4ff8b-106">Attribute</span></span> | <span data-ttu-id="4ff8b-107">必需</span><span class="sxs-lookup"><span data-stu-id="4ff8b-107">Required</span></span> | <span data-ttu-id="4ff8b-108">说明</span><span class="sxs-lookup"><span data-stu-id="4ff8b-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="4ff8b-109">resid</span><span class="sxs-lookup"><span data-stu-id="4ff8b-109">resid</span></span>     | <span data-ttu-id="4ff8b-110">是</span><span class="sxs-lookup"><span data-stu-id="4ff8b-110">Yes</span></span>      | <span data-ttu-id="4ff8b-111">清单的 &lt;Resources&gt; 部分中所定义的 URL 资源的名称。</span><span class="sxs-lookup"><span data-stu-id="4ff8b-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="4ff8b-112">子元素</span><span class="sxs-lookup"><span data-stu-id="4ff8b-112">Child elements</span></span>

<span data-ttu-id="4ff8b-113">无</span><span class="sxs-lookup"><span data-stu-id="4ff8b-113">None</span></span>

## <a name="example"></a><span data-ttu-id="4ff8b-114">示例</span><span class="sxs-lookup"><span data-stu-id="4ff8b-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
