---
title: 清单文件中的 Page 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433731"
---
# <a name="page-element"></a><span data-ttu-id="a31d3-102">Page 元素</span><span class="sxs-lookup"><span data-stu-id="a31d3-102">Page element</span></span>

<span data-ttu-id="a31d3-103">定义 Excel 中的自定义函数所使用的 HTML 页面设置。</span><span class="sxs-lookup"><span data-stu-id="a31d3-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a31d3-104">属性</span><span class="sxs-lookup"><span data-stu-id="a31d3-104">Attributes</span></span>

<span data-ttu-id="a31d3-105">无</span><span class="sxs-lookup"><span data-stu-id="a31d3-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="a31d3-106">子元素</span><span class="sxs-lookup"><span data-stu-id="a31d3-106">Child elements</span></span>

|  <span data-ttu-id="a31d3-107">元素</span><span class="sxs-lookup"><span data-stu-id="a31d3-107">Element</span></span>  |  <span data-ttu-id="a31d3-108">必需</span><span class="sxs-lookup"><span data-stu-id="a31d3-108">Required</span></span>  |  <span data-ttu-id="a31d3-109">说明</span><span class="sxs-lookup"><span data-stu-id="a31d3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a31d3-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a31d3-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="a31d3-111">是</span><span class="sxs-lookup"><span data-stu-id="a31d3-111">Yes</span></span>  | <span data-ttu-id="a31d3-112">包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="a31d3-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="a31d3-113">示例</span><span class="sxs-lookup"><span data-stu-id="a31d3-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
