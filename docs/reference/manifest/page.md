---
title: 清单文件中的 Page 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f85cc3a834f628a7390f3b96faa596145c7d331a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452072"
---
# <a name="page-element"></a><span data-ttu-id="b88f9-102">Page 元素</span><span class="sxs-lookup"><span data-stu-id="b88f9-102">Page element</span></span>

<span data-ttu-id="b88f9-103">定义 Excel 中的自定义函数所使用的 HTML 页面设置。</span><span class="sxs-lookup"><span data-stu-id="b88f9-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="b88f9-104">属性</span><span class="sxs-lookup"><span data-stu-id="b88f9-104">Attributes</span></span>

<span data-ttu-id="b88f9-105">无</span><span class="sxs-lookup"><span data-stu-id="b88f9-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="b88f9-106">子元素</span><span class="sxs-lookup"><span data-stu-id="b88f9-106">Child elements</span></span>

|  <span data-ttu-id="b88f9-107">元素</span><span class="sxs-lookup"><span data-stu-id="b88f9-107">Element</span></span>  |  <span data-ttu-id="b88f9-108">必需</span><span class="sxs-lookup"><span data-stu-id="b88f9-108">Required</span></span>  |  <span data-ttu-id="b88f9-109">说明</span><span class="sxs-lookup"><span data-stu-id="b88f9-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b88f9-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="b88f9-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="b88f9-111">是</span><span class="sxs-lookup"><span data-stu-id="b88f9-111">Yes</span></span>  | <span data-ttu-id="b88f9-112">包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="b88f9-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="b88f9-113">示例</span><span class="sxs-lookup"><span data-stu-id="b88f9-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
