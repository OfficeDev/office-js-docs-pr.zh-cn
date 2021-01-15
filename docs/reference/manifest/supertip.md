---
title: 清单文件中的 Supertip 元素
description: Supertip 元素定义一个丰富的工具提示 (标题和说明) 。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771296"
---
# <a name="supertip"></a><span data-ttu-id="0cdbb-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="0cdbb-103">Supertip</span></span>

<span data-ttu-id="0cdbb-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0cdbb-106">子元素</span><span class="sxs-lookup"><span data-stu-id="0cdbb-106">Child elements</span></span>

|  <span data-ttu-id="0cdbb-107">元素</span><span class="sxs-lookup"><span data-stu-id="0cdbb-107">Element</span></span> |  <span data-ttu-id="0cdbb-108">必需</span><span class="sxs-lookup"><span data-stu-id="0cdbb-108">Required</span></span>  |  <span data-ttu-id="0cdbb-109">说明</span><span class="sxs-lookup"><span data-stu-id="0cdbb-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="0cdbb-110">标题</span><span class="sxs-lookup"><span data-stu-id="0cdbb-110">Title</span></span>](#title) | <span data-ttu-id="0cdbb-111">是</span><span class="sxs-lookup"><span data-stu-id="0cdbb-111">Yes</span></span> | <span data-ttu-id="0cdbb-112">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="0cdbb-113">说明</span><span class="sxs-lookup"><span data-stu-id="0cdbb-113">Description</span></span>](#description) | <span data-ttu-id="0cdbb-114">是</span><span class="sxs-lookup"><span data-stu-id="0cdbb-114">Yes</span></span> | <span data-ttu-id="0cdbb-115">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-115">The description for the supertip.</span></span><br><span data-ttu-id="0cdbb-116">**注意**： (Outlook) 仅支持 Windows 和 Mac 客户端。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="0cdbb-117">标题</span><span class="sxs-lookup"><span data-stu-id="0cdbb-117">Title</span></span>

<span data-ttu-id="0cdbb-118">必需。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-118">Required.</span></span> <span data-ttu-id="0cdbb-119">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-119">The text for the supertip.</span></span> <span data-ttu-id="0cdbb-120">**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="0cdbb-120">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="0cdbb-121">说明</span><span class="sxs-lookup"><span data-stu-id="0cdbb-121">Description</span></span>

<span data-ttu-id="0cdbb-122">必需。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-122">Required.</span></span> <span data-ttu-id="0cdbb-123">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-123">The description for the supertip.</span></span> <span data-ttu-id="0cdbb-124">**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 LongStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="0cdbb-124">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="0cdbb-125">对于 Outlook，只有 Windows 和 Mac 客户端支持 **Description** 元素。</span><span class="sxs-lookup"><span data-stu-id="0cdbb-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="0cdbb-126">示例</span><span class="sxs-lookup"><span data-stu-id="0cdbb-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
