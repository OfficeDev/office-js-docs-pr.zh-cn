---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在 Excel 中使用的命名空间。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771255"
---
# <a name="namespace-element"></a><span data-ttu-id="a738b-103">Namespace 元素</span><span class="sxs-lookup"><span data-stu-id="a738b-103">Namespace element</span></span>

<span data-ttu-id="a738b-104">定义 Excel 中的自定义函数使用的命名空间。</span><span class="sxs-lookup"><span data-stu-id="a738b-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a738b-105">属性</span><span class="sxs-lookup"><span data-stu-id="a738b-105">Attributes</span></span>

|  <span data-ttu-id="a738b-106">属性</span><span class="sxs-lookup"><span data-stu-id="a738b-106">Attribute</span></span>  |  <span data-ttu-id="a738b-107">必需</span><span class="sxs-lookup"><span data-stu-id="a738b-107">Required</span></span>  |  <span data-ttu-id="a738b-108">说明</span><span class="sxs-lookup"><span data-stu-id="a738b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a738b-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="a738b-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="a738b-110">否</span><span class="sxs-lookup"><span data-stu-id="a738b-110">No</span></span>  | <span data-ttu-id="a738b-111">应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。</span><span class="sxs-lookup"><span data-stu-id="a738b-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> <span data-ttu-id="a738b-112">不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="a738b-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="a738b-113">子元素</span><span class="sxs-lookup"><span data-stu-id="a738b-113">Child elements</span></span>

<span data-ttu-id="a738b-114">无</span><span class="sxs-lookup"><span data-stu-id="a738b-114">None</span></span>

## <a name="example"></a><span data-ttu-id="a738b-115">示例</span><span class="sxs-lookup"><span data-stu-id="a738b-115">Example</span></span>

```xml
<Namespace resid="namespace" />
```
