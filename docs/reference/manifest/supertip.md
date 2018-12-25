---
title: 清单文件中的 Supertip 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: bae997eda8e1055c5be76382456ba83acca7b91c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433668"
---
# <a name="supertip"></a><span data-ttu-id="b979b-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="b979b-102">Supertip</span></span>

<span data-ttu-id="b979b-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="b979b-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b979b-105">子元素</span><span class="sxs-lookup"><span data-stu-id="b979b-105">Child elements</span></span>

|  <span data-ttu-id="b979b-106">元素</span><span class="sxs-lookup"><span data-stu-id="b979b-106">Element</span></span> |  <span data-ttu-id="b979b-107">必需</span><span class="sxs-lookup"><span data-stu-id="b979b-107">Required</span></span>  |  <span data-ttu-id="b979b-108">说明</span><span class="sxs-lookup"><span data-stu-id="b979b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b979b-109">Title</span><span class="sxs-lookup"><span data-stu-id="b979b-109">Title</span></span>](#title)        | <span data-ttu-id="b979b-110">是</span><span class="sxs-lookup"><span data-stu-id="b979b-110">Yes</span></span> |   <span data-ttu-id="b979b-111">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="b979b-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="b979b-112">说明</span><span class="sxs-lookup"><span data-stu-id="b979b-112">Description</span></span>](#description)  | <span data-ttu-id="b979b-113">是</span><span class="sxs-lookup"><span data-stu-id="b979b-113">Yes</span></span> |  <span data-ttu-id="b979b-114">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="b979b-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="b979b-115">标题</span><span class="sxs-lookup"><span data-stu-id="b979b-115">Title</span></span>

<span data-ttu-id="b979b-p102">必需。SuperTip 的文本。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="b979b-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="b979b-119">说明</span><span class="sxs-lookup"><span data-stu-id="b979b-119">Description</span></span>

<span data-ttu-id="b979b-p103">必需。SuperTip 的描述。 **resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="b979b-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="b979b-123">示例</span><span class="sxs-lookup"><span data-stu-id="b979b-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
