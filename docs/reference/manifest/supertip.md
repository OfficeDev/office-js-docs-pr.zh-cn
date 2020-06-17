---
title: 清单文件中的 Supertip 元素
description: Supertip 元素定义了一个丰富的工具提示（标题和说明）。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 8061c9dcd7903db0f1265084498d6c86654e1dfa
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608717"
---
# <a name="supertip"></a><span data-ttu-id="c8a92-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="c8a92-103">Supertip</span></span>

<span data-ttu-id="c8a92-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="c8a92-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c8a92-106">子元素</span><span class="sxs-lookup"><span data-stu-id="c8a92-106">Child elements</span></span>

|  <span data-ttu-id="c8a92-107">元素</span><span class="sxs-lookup"><span data-stu-id="c8a92-107">Element</span></span> |  <span data-ttu-id="c8a92-108">必需</span><span class="sxs-lookup"><span data-stu-id="c8a92-108">Required</span></span>  |  <span data-ttu-id="c8a92-109">Description</span><span class="sxs-lookup"><span data-stu-id="c8a92-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c8a92-110">标题</span><span class="sxs-lookup"><span data-stu-id="c8a92-110">Title</span></span>](#title) | <span data-ttu-id="c8a92-111">是</span><span class="sxs-lookup"><span data-stu-id="c8a92-111">Yes</span></span> | <span data-ttu-id="c8a92-112">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="c8a92-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="c8a92-113">说明</span><span class="sxs-lookup"><span data-stu-id="c8a92-113">Description</span></span>](#description) | <span data-ttu-id="c8a92-114">是</span><span class="sxs-lookup"><span data-stu-id="c8a92-114">Yes</span></span> | <span data-ttu-id="c8a92-115">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="c8a92-115">The description for the supertip.</span></span><br><span data-ttu-id="c8a92-116">**注意**：（Outlook）仅支持 Windows 和 Mac 客户端。</span><span class="sxs-lookup"><span data-stu-id="c8a92-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="c8a92-117">Title</span><span class="sxs-lookup"><span data-stu-id="c8a92-117">Title</span></span>

<span data-ttu-id="c8a92-118">必填。</span><span class="sxs-lookup"><span data-stu-id="c8a92-118">Required.</span></span> <span data-ttu-id="c8a92-119">SuperTip 的文本。</span><span class="sxs-lookup"><span data-stu-id="c8a92-119">The text for the supertip.</span></span> <span data-ttu-id="c8a92-120">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="c8a92-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="c8a92-121">说明</span><span class="sxs-lookup"><span data-stu-id="c8a92-121">Description</span></span>

<span data-ttu-id="c8a92-122">必需。</span><span class="sxs-lookup"><span data-stu-id="c8a92-122">Required.</span></span> <span data-ttu-id="c8a92-123">SuperTip 的描述。</span><span class="sxs-lookup"><span data-stu-id="c8a92-123">The description for the supertip.</span></span> <span data-ttu-id="c8a92-124">**Resid**属性必须设置为[Resources](resources.md)元素中的**LongStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="c8a92-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="c8a92-125">对于 Outlook，只有 Windows 和 Mac 客户端支持**Description**元素。</span><span class="sxs-lookup"><span data-stu-id="c8a92-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="c8a92-126">示例</span><span class="sxs-lookup"><span data-stu-id="c8a92-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
