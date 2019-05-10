---
title: 清单文件中的 Supertip 元素
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659652"
---
# <a name="supertip"></a><span data-ttu-id="30592-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="30592-102">Supertip</span></span>

<span data-ttu-id="30592-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="30592-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="30592-105">子元素</span><span class="sxs-lookup"><span data-stu-id="30592-105">Child elements</span></span>

|  <span data-ttu-id="30592-106">元素</span><span class="sxs-lookup"><span data-stu-id="30592-106">Element</span></span> |  <span data-ttu-id="30592-107">必需</span><span class="sxs-lookup"><span data-stu-id="30592-107">Required</span></span>  |  <span data-ttu-id="30592-108">说明</span><span class="sxs-lookup"><span data-stu-id="30592-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="30592-109">标题</span><span class="sxs-lookup"><span data-stu-id="30592-109">Title</span></span>](#title) | <span data-ttu-id="30592-110">是</span><span class="sxs-lookup"><span data-stu-id="30592-110">Yes</span></span> | <span data-ttu-id="30592-111">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="30592-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="30592-112">说明</span><span class="sxs-lookup"><span data-stu-id="30592-112">Description</span></span>](#description) | <span data-ttu-id="30592-113">是</span><span class="sxs-lookup"><span data-stu-id="30592-113">Yes</span></span> | <span data-ttu-id="30592-114">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="30592-114">The description for the supertip.</span></span><br><span data-ttu-id="30592-115">**注意**: (Outlook) 仅支持 Windows 和 Mac 客户端。</span><span class="sxs-lookup"><span data-stu-id="30592-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="30592-116">Title</span><span class="sxs-lookup"><span data-stu-id="30592-116">Title</span></span>

<span data-ttu-id="30592-p102">必需。SuperTip 的文本。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="30592-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="30592-120">说明</span><span class="sxs-lookup"><span data-stu-id="30592-120">Description</span></span>

<span data-ttu-id="30592-p103">必需。SuperTip 的描述。 **resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="30592-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="30592-124">对于 Outlook, 只有 Windows 和 Mac 客户端支持**Description**元素。</span><span class="sxs-lookup"><span data-stu-id="30592-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="30592-125">示例</span><span class="sxs-lookup"><span data-stu-id="30592-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
