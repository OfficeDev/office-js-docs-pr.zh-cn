---
title: 清单文件中的 Namespace 元素
description: Namespace 元素定义自定义函数在 Excel 中使用的命名空间。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612232"
---
# <a name="namespace-element"></a><span data-ttu-id="a964b-103">Namespace 元素</span><span class="sxs-lookup"><span data-stu-id="a964b-103">Namespace element</span></span>

<span data-ttu-id="a964b-104">定义 Excel 中的自定义函数使用的命名空间。</span><span class="sxs-lookup"><span data-stu-id="a964b-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a964b-105">属性</span><span class="sxs-lookup"><span data-stu-id="a964b-105">Attributes</span></span>

|  <span data-ttu-id="a964b-106">属性</span><span class="sxs-lookup"><span data-stu-id="a964b-106">Attribute</span></span>  |  <span data-ttu-id="a964b-107">必需</span><span class="sxs-lookup"><span data-stu-id="a964b-107">Required</span></span>  |  <span data-ttu-id="a964b-108">Description</span><span class="sxs-lookup"><span data-stu-id="a964b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a964b-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="a964b-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="a964b-110">否</span><span class="sxs-lookup"><span data-stu-id="a964b-110">No</span></span>  | <span data-ttu-id="a964b-111">应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。</span><span class="sxs-lookup"><span data-stu-id="a964b-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="a964b-112">子元素</span><span class="sxs-lookup"><span data-stu-id="a964b-112">Child elements</span></span>

<span data-ttu-id="a964b-113">无</span><span class="sxs-lookup"><span data-stu-id="a964b-113">None</span></span>

## <a name="example"></a><span data-ttu-id="a964b-114">示例</span><span class="sxs-lookup"><span data-stu-id="a964b-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
