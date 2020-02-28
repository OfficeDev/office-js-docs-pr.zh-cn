---
title: 清单文件中的 Supertip 元素
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325232"
---
# <a name="supertip"></a><span data-ttu-id="91bb3-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="91bb3-102">Supertip</span></span>

<span data-ttu-id="91bb3-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="91bb3-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="91bb3-105">子元素</span><span class="sxs-lookup"><span data-stu-id="91bb3-105">Child elements</span></span>

|  <span data-ttu-id="91bb3-106">元素</span><span class="sxs-lookup"><span data-stu-id="91bb3-106">Element</span></span> |  <span data-ttu-id="91bb3-107">必需</span><span class="sxs-lookup"><span data-stu-id="91bb3-107">Required</span></span>  |  <span data-ttu-id="91bb3-108">说明</span><span class="sxs-lookup"><span data-stu-id="91bb3-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="91bb3-109">标题</span><span class="sxs-lookup"><span data-stu-id="91bb3-109">Title</span></span>](#title) | <span data-ttu-id="91bb3-110">是</span><span class="sxs-lookup"><span data-stu-id="91bb3-110">Yes</span></span> | <span data-ttu-id="91bb3-111">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="91bb3-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="91bb3-112">说明</span><span class="sxs-lookup"><span data-stu-id="91bb3-112">Description</span></span>](#description) | <span data-ttu-id="91bb3-113">是</span><span class="sxs-lookup"><span data-stu-id="91bb3-113">Yes</span></span> | <span data-ttu-id="91bb3-114">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="91bb3-114">The description for the supertip.</span></span><br><span data-ttu-id="91bb3-115">**注意**：（Outlook）仅支持 Windows 和 Mac 客户端。</span><span class="sxs-lookup"><span data-stu-id="91bb3-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="91bb3-116">Title</span><span class="sxs-lookup"><span data-stu-id="91bb3-116">Title</span></span>

<span data-ttu-id="91bb3-117">必填。</span><span class="sxs-lookup"><span data-stu-id="91bb3-117">Required.</span></span> <span data-ttu-id="91bb3-118">SuperTip 的文本。</span><span class="sxs-lookup"><span data-stu-id="91bb3-118">The text for the supertip.</span></span> <span data-ttu-id="91bb3-119">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="91bb3-119">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="91bb3-120">说明</span><span class="sxs-lookup"><span data-stu-id="91bb3-120">Description</span></span>

<span data-ttu-id="91bb3-121">必需。</span><span class="sxs-lookup"><span data-stu-id="91bb3-121">Required.</span></span> <span data-ttu-id="91bb3-122">SuperTip 的描述。</span><span class="sxs-lookup"><span data-stu-id="91bb3-122">The description for the supertip.</span></span> <span data-ttu-id="91bb3-123">**Resid**属性必须设置为[Resources](resources.md)元素中的**LongStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="91bb3-123">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="91bb3-124">对于 Outlook，只有 Windows 和 Mac 客户端支持**Description**元素。</span><span class="sxs-lookup"><span data-stu-id="91bb3-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="91bb3-125">示例</span><span class="sxs-lookup"><span data-stu-id="91bb3-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
