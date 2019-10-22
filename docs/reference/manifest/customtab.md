---
title: 清单文件中的 CustomTab 元素
description: ''
ms.date: 04/29/2019
localization_priority: Normal
ms.openlocfilehash: 4fa7dd86736b5ab421be5653f2e256a6b84fb480
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/21/2019
ms.locfileid: "33517392"
---
# <a name="customtab-element"></a><span data-ttu-id="1d93e-102">CustomTab 元素</span><span class="sxs-lookup"><span data-stu-id="1d93e-102">CustomTab element</span></span>

<span data-ttu-id="1d93e-p101">在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。这可以位于默认的选项卡（“**开始**”、“**消息**”或“**会议**”）上，或位于由外接程序定义的自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="1d93e-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="1d93e-p102">在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="1d93e-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="1d93e-108">**id** 属性在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="1d93e-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1d93e-109">子元素</span><span class="sxs-lookup"><span data-stu-id="1d93e-109">Child elements</span></span>

|  <span data-ttu-id="1d93e-110">元素</span><span class="sxs-lookup"><span data-stu-id="1d93e-110">Element</span></span> |  <span data-ttu-id="1d93e-111">必需</span><span class="sxs-lookup"><span data-stu-id="1d93e-111">Required</span></span>  |  <span data-ttu-id="1d93e-112">说明</span><span class="sxs-lookup"><span data-stu-id="1d93e-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1d93e-113">Group</span><span class="sxs-lookup"><span data-stu-id="1d93e-113">Group</span></span>](group.md)      | <span data-ttu-id="1d93e-114">是</span><span class="sxs-lookup"><span data-stu-id="1d93e-114">Yes</span></span> |  <span data-ttu-id="1d93e-115">定义一组命令。</span><span class="sxs-lookup"><span data-stu-id="1d93e-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="1d93e-116">Label</span><span class="sxs-lookup"><span data-stu-id="1d93e-116">Label</span></span>](#label-tab)      | <span data-ttu-id="1d93e-117">是</span><span class="sxs-lookup"><span data-stu-id="1d93e-117">Yes</span></span> |  <span data-ttu-id="1d93e-118">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="1d93e-118">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="1d93e-119">组</span><span class="sxs-lookup"><span data-stu-id="1d93e-119">Group</span></span>

<span data-ttu-id="1d93e-p103">必需。查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="1d93e-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="1d93e-122">标签（选项卡）</span><span class="sxs-lookup"><span data-stu-id="1d93e-122">Label (Tab)</span></span>

<span data-ttu-id="1d93e-p104">必需。自定义选项卡的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="1d93e-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="1d93e-125">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="1d93e-125">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
