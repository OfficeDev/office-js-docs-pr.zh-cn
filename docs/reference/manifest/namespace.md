---
title: 清单文件中的 Namespace 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: faf77fe8b6bddc734f1b47eb544ffe7e1e7c4aaa
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452100"
---
# <a name="namespace-element"></a><span data-ttu-id="8c02f-102">Namespace 元素</span><span class="sxs-lookup"><span data-stu-id="8c02f-102">Namespace element</span></span>

<span data-ttu-id="8c02f-103">定义 Excel 中的自定义函数使用的命名空间。</span><span class="sxs-lookup"><span data-stu-id="8c02f-103">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="8c02f-104">属性</span><span class="sxs-lookup"><span data-stu-id="8c02f-104">Attributes</span></span>

|  <span data-ttu-id="8c02f-105">属性</span><span class="sxs-lookup"><span data-stu-id="8c02f-105">Attribute</span></span>  |  <span data-ttu-id="8c02f-106">必需</span><span class="sxs-lookup"><span data-stu-id="8c02f-106">Required</span></span>  |  <span data-ttu-id="8c02f-107">说明</span><span class="sxs-lookup"><span data-stu-id="8c02f-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="8c02f-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="8c02f-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="8c02f-109">是</span><span class="sxs-lookup"><span data-stu-id="8c02f-109">Yes</span></span>  | <span data-ttu-id="8c02f-110">应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。</span><span class="sxs-lookup"><span data-stu-id="8c02f-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="8c02f-111">子元素</span><span class="sxs-lookup"><span data-stu-id="8c02f-111">Child elements</span></span>

<span data-ttu-id="8c02f-112">无</span><span class="sxs-lookup"><span data-stu-id="8c02f-112">None</span></span>

## <a name="example"></a><span data-ttu-id="8c02f-113">示例</span><span class="sxs-lookup"><span data-stu-id="8c02f-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
