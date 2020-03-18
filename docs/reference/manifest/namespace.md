---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在 Excel 中使用的命名空间。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45fd0caa039fdeb885cba4b739750fbd8b642252
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718053"
---
# <a name="namespace-element"></a><span data-ttu-id="70500-103">Namespace 元素</span><span class="sxs-lookup"><span data-stu-id="70500-103">Namespace element</span></span>

<span data-ttu-id="70500-104">定义 Excel 中的自定义函数使用的命名空间。</span><span class="sxs-lookup"><span data-stu-id="70500-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="70500-105">属性</span><span class="sxs-lookup"><span data-stu-id="70500-105">Attributes</span></span>

|  <span data-ttu-id="70500-106">属性</span><span class="sxs-lookup"><span data-stu-id="70500-106">Attribute</span></span>  |  <span data-ttu-id="70500-107">必需</span><span class="sxs-lookup"><span data-stu-id="70500-107">Required</span></span>  |  <span data-ttu-id="70500-108">说明</span><span class="sxs-lookup"><span data-stu-id="70500-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="70500-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="70500-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="70500-110">是</span><span class="sxs-lookup"><span data-stu-id="70500-110">Yes</span></span>  | <span data-ttu-id="70500-111">应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。</span><span class="sxs-lookup"><span data-stu-id="70500-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="70500-112">子元素</span><span class="sxs-lookup"><span data-stu-id="70500-112">Child elements</span></span>

<span data-ttu-id="70500-113">无</span><span class="sxs-lookup"><span data-stu-id="70500-113">None</span></span>

## <a name="example"></a><span data-ttu-id="70500-114">示例</span><span class="sxs-lookup"><span data-stu-id="70500-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
