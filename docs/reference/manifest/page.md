---
title: 清单文件中的 Page 元素
description: Page 元素定义了自定义函数在 Excel 中使用的 HTML 页面设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720482"
---
# <a name="page-element"></a><span data-ttu-id="48217-103">Page 元素</span><span class="sxs-lookup"><span data-stu-id="48217-103">Page element</span></span>

<span data-ttu-id="48217-104">定义 Excel 中的自定义函数所使用的 HTML 页面设置。</span><span class="sxs-lookup"><span data-stu-id="48217-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="48217-105">属性</span><span class="sxs-lookup"><span data-stu-id="48217-105">Attributes</span></span>

<span data-ttu-id="48217-106">无</span><span class="sxs-lookup"><span data-stu-id="48217-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="48217-107">子元素</span><span class="sxs-lookup"><span data-stu-id="48217-107">Child elements</span></span>

|  <span data-ttu-id="48217-108">元素</span><span class="sxs-lookup"><span data-stu-id="48217-108">Element</span></span>  |  <span data-ttu-id="48217-109">必需</span><span class="sxs-lookup"><span data-stu-id="48217-109">Required</span></span>  |  <span data-ttu-id="48217-110">说明</span><span class="sxs-lookup"><span data-stu-id="48217-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="48217-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="48217-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="48217-112">是</span><span class="sxs-lookup"><span data-stu-id="48217-112">Yes</span></span>  | <span data-ttu-id="48217-113">包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="48217-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="48217-114">示例</span><span class="sxs-lookup"><span data-stu-id="48217-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
