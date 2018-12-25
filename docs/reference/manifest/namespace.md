---
title: 清单文件中的 Namespace 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432324"
---
# <a name="namespace-element"></a><span data-ttu-id="eef38-102">Namespace 元素</span><span class="sxs-lookup"><span data-stu-id="eef38-102">Namespace element</span></span>

<span data-ttu-id="eef38-103">定义 Excel 中的自定义函数使用的命名空间。</span><span class="sxs-lookup"><span data-stu-id="eef38-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="eef38-104">属性</span><span class="sxs-lookup"><span data-stu-id="eef38-104">Attributes</span></span>

|  <span data-ttu-id="eef38-105">属性</span><span class="sxs-lookup"><span data-stu-id="eef38-105">Attribute</span></span>  |  <span data-ttu-id="eef38-106">必需</span><span class="sxs-lookup"><span data-stu-id="eef38-106">Required</span></span>  |  <span data-ttu-id="eef38-107">说明</span><span class="sxs-lookup"><span data-stu-id="eef38-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="eef38-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="eef38-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="eef38-109">是</span><span class="sxs-lookup"><span data-stu-id="eef38-109">Yes</span></span>  | <span data-ttu-id="eef38-110">应与 [Resources](resources.md) 元素中指定的自定义函数的 ShortStrings 标题匹配。</span><span class="sxs-lookup"><span data-stu-id="eef38-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="eef38-111">子元素</span><span class="sxs-lookup"><span data-stu-id="eef38-111">Child elements</span></span>

<span data-ttu-id="eef38-112">无</span><span class="sxs-lookup"><span data-stu-id="eef38-112">None</span></span>

## <a name="example"></a><span data-ttu-id="eef38-113">示例</span><span class="sxs-lookup"><span data-stu-id="eef38-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
