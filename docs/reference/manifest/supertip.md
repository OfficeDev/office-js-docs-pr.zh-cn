---
title: 清单文件中的 Supertip 元素
description: Supertip 元素定义了一个丰富的工具提示（标题和说明）。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: cf88473b72979c839e5d55f44938fda19be24084
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720349"
---
# <a name="supertip"></a><span data-ttu-id="614e8-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="614e8-103">Supertip</span></span>

<span data-ttu-id="614e8-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="614e8-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="614e8-106">子元素</span><span class="sxs-lookup"><span data-stu-id="614e8-106">Child elements</span></span>

|  <span data-ttu-id="614e8-107">元素</span><span class="sxs-lookup"><span data-stu-id="614e8-107">Element</span></span> |  <span data-ttu-id="614e8-108">必需</span><span class="sxs-lookup"><span data-stu-id="614e8-108">Required</span></span>  |  <span data-ttu-id="614e8-109">说明</span><span class="sxs-lookup"><span data-stu-id="614e8-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="614e8-110">标题</span><span class="sxs-lookup"><span data-stu-id="614e8-110">Title</span></span>](#title) | <span data-ttu-id="614e8-111">是</span><span class="sxs-lookup"><span data-stu-id="614e8-111">Yes</span></span> | <span data-ttu-id="614e8-112">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="614e8-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="614e8-113">说明</span><span class="sxs-lookup"><span data-stu-id="614e8-113">Description</span></span>](#description) | <span data-ttu-id="614e8-114">是</span><span class="sxs-lookup"><span data-stu-id="614e8-114">Yes</span></span> | <span data-ttu-id="614e8-115">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="614e8-115">The description for the supertip.</span></span><br><span data-ttu-id="614e8-116">**注意**：（Outlook）仅支持 Windows 和 Mac 客户端。</span><span class="sxs-lookup"><span data-stu-id="614e8-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="614e8-117">Title</span><span class="sxs-lookup"><span data-stu-id="614e8-117">Title</span></span>

<span data-ttu-id="614e8-118">必需。</span><span class="sxs-lookup"><span data-stu-id="614e8-118">Required.</span></span> <span data-ttu-id="614e8-119">SuperTip 的文本。</span><span class="sxs-lookup"><span data-stu-id="614e8-119">The text for the supertip.</span></span> <span data-ttu-id="614e8-120">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="614e8-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="614e8-121">说明</span><span class="sxs-lookup"><span data-stu-id="614e8-121">Description</span></span>

<span data-ttu-id="614e8-122">必需。</span><span class="sxs-lookup"><span data-stu-id="614e8-122">Required.</span></span> <span data-ttu-id="614e8-123">SuperTip 的描述。</span><span class="sxs-lookup"><span data-stu-id="614e8-123">The description for the supertip.</span></span> <span data-ttu-id="614e8-124">**Resid**属性必须设置为[Resources](resources.md)元素中的**LongStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="614e8-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="614e8-125">对于 Outlook，只有 Windows 和 Mac 客户端支持**Description**元素。</span><span class="sxs-lookup"><span data-stu-id="614e8-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="614e8-126">示例</span><span class="sxs-lookup"><span data-stu-id="614e8-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
