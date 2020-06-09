---
title: 清单文件中的 Page 元素
description: Page 元素定义了自定义函数在 Excel 中使用的 HTML 页面设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611496"
---
# <a name="page-element"></a><span data-ttu-id="6d035-103">Page 元素</span><span class="sxs-lookup"><span data-stu-id="6d035-103">Page element</span></span>

<span data-ttu-id="6d035-104">定义 Excel 中的自定义函数所使用的 HTML 页面设置。</span><span class="sxs-lookup"><span data-stu-id="6d035-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="6d035-105">属性</span><span class="sxs-lookup"><span data-stu-id="6d035-105">Attributes</span></span>

<span data-ttu-id="6d035-106">无</span><span class="sxs-lookup"><span data-stu-id="6d035-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="6d035-107">子元素</span><span class="sxs-lookup"><span data-stu-id="6d035-107">Child elements</span></span>

|  <span data-ttu-id="6d035-108">元素</span><span class="sxs-lookup"><span data-stu-id="6d035-108">Element</span></span>  |  <span data-ttu-id="6d035-109">必需</span><span class="sxs-lookup"><span data-stu-id="6d035-109">Required</span></span>  |  <span data-ttu-id="6d035-110">Description</span><span class="sxs-lookup"><span data-stu-id="6d035-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6d035-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6d035-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="6d035-112">是</span><span class="sxs-lookup"><span data-stu-id="6d035-112">Yes</span></span>  | <span data-ttu-id="6d035-113">包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="6d035-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="6d035-114">示例</span><span class="sxs-lookup"><span data-stu-id="6d035-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
