---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: a81b64a17eeeb463d55024e189b09048b2eb96ac
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612302"
---
# <a name="customtab-element"></a><span data-ttu-id="9cf01-103">CustomTab 元素</span><span class="sxs-lookup"><span data-stu-id="9cf01-103">CustomTab element</span></span>

<span data-ttu-id="9cf01-104">在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。</span><span class="sxs-lookup"><span data-stu-id="9cf01-104">On the ribbon, you specify which tab and group for their add-in commands.</span></span> <span data-ttu-id="9cf01-105">这可能位于默认选项卡（“主页”\*\*\*\*、“邮件”\*\*\*\* 或“会议”\*\*\*\*）上，或位于外接程序定义的自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="9cf01-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="9cf01-p102">在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="9cf01-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="9cf01-109">**Id**属性在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="9cf01-109">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9cf01-110">在 Mac 上的 Outlook 中，该元素不可用， `CustomTab` 因此您必须改用[OfficeTab](officetab.md) 。</span><span class="sxs-lookup"><span data-stu-id="9cf01-110">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9cf01-111">子元素</span><span class="sxs-lookup"><span data-stu-id="9cf01-111">Child elements</span></span>

|  <span data-ttu-id="9cf01-112">元素</span><span class="sxs-lookup"><span data-stu-id="9cf01-112">Element</span></span> |  <span data-ttu-id="9cf01-113">必需</span><span class="sxs-lookup"><span data-stu-id="9cf01-113">Required</span></span>  |  <span data-ttu-id="9cf01-114">说明</span><span class="sxs-lookup"><span data-stu-id="9cf01-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9cf01-115">Group</span><span class="sxs-lookup"><span data-stu-id="9cf01-115">Group</span></span>](group.md)      | <span data-ttu-id="9cf01-116">是</span><span class="sxs-lookup"><span data-stu-id="9cf01-116">Yes</span></span> |  <span data-ttu-id="9cf01-117">定义一组命令。</span><span class="sxs-lookup"><span data-stu-id="9cf01-117">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="9cf01-118">Label</span><span class="sxs-lookup"><span data-stu-id="9cf01-118">Label</span></span>](#label-tab)      | <span data-ttu-id="9cf01-119">是</span><span class="sxs-lookup"><span data-stu-id="9cf01-119">Yes</span></span> |  <span data-ttu-id="9cf01-120">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="9cf01-120">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="9cf01-121">组</span><span class="sxs-lookup"><span data-stu-id="9cf01-121">Group</span></span>

<span data-ttu-id="9cf01-p103">必需。查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="9cf01-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="9cf01-124">标签（选项卡）</span><span class="sxs-lookup"><span data-stu-id="9cf01-124">Label (Tab)</span></span>

<span data-ttu-id="9cf01-125">必填。</span><span class="sxs-lookup"><span data-stu-id="9cf01-125">Required.</span></span> <span data-ttu-id="9cf01-126">自定义选项卡的标签。**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="9cf01-126">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="9cf01-127">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="9cf01-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
