---
title: 清单文件中的 GetStarted 元素
description: 提供在 Word、Excel、PowerPoint 和 OneNote 主机中安装此外接程序时显示的标注所使用的信息。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7b4eb0e0413eff97751eac01c3dc6e5a452722aa
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718186"
---
# <a name="getstarted-element"></a><span data-ttu-id="77ef7-103">GetStarted 元素</span><span class="sxs-lookup"><span data-stu-id="77ef7-103">GetStarted element</span></span>

<span data-ttu-id="77ef7-p101">提供在 Word、Excel、PowerPoint 和 OneNote 主机中安装此外接程序时显示的标注所使用的信息。**GetStarted** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="77ef7-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="77ef7-106">子元素</span><span class="sxs-lookup"><span data-stu-id="77ef7-106">Child elements</span></span>

| <span data-ttu-id="77ef7-107">元素</span><span class="sxs-lookup"><span data-stu-id="77ef7-107">Element</span></span>                       | <span data-ttu-id="77ef7-108">必需</span><span class="sxs-lookup"><span data-stu-id="77ef7-108">Required</span></span> | <span data-ttu-id="77ef7-109">说明</span><span class="sxs-lookup"><span data-stu-id="77ef7-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="77ef7-110">标题</span><span class="sxs-lookup"><span data-stu-id="77ef7-110">Title</span></span>](#title)               | <span data-ttu-id="77ef7-111">是</span><span class="sxs-lookup"><span data-stu-id="77ef7-111">Yes</span></span>      | <span data-ttu-id="77ef7-112">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="77ef7-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="77ef7-113">说明</span><span class="sxs-lookup"><span data-stu-id="77ef7-113">Description</span></span>](#description)   | <span data-ttu-id="77ef7-114">是</span><span class="sxs-lookup"><span data-stu-id="77ef7-114">Yes</span></span>      | <span data-ttu-id="77ef7-115">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="77ef7-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="77ef7-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="77ef7-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="77ef7-117">是</span><span class="sxs-lookup"><span data-stu-id="77ef7-117">Yes</span></span>       | <span data-ttu-id="77ef7-118">指向详细说明外接程序的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="77ef7-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="77ef7-119">标题</span><span class="sxs-lookup"><span data-stu-id="77ef7-119">Title</span></span> 

<span data-ttu-id="77ef7-p102">必需。 用于标注顶部的标题。 **resid** 属性引用 **Resources** 分区的 [ShortStrings](resources.md) 元素中的有效 ID。</span><span class="sxs-lookup"><span data-stu-id="77ef7-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="77ef7-123">说明</span><span class="sxs-lookup"><span data-stu-id="77ef7-123">Description</span></span>

<span data-ttu-id="77ef7-p103">必需。 标注的说明/正文内容。 **resid** 属性引用 **Resources** 分区的 [LongStrings](resources.md) 元素中的有效 ID。</span><span class="sxs-lookup"><span data-stu-id="77ef7-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="77ef7-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="77ef7-127">LearnMoreUrl</span></span>

<span data-ttu-id="77ef7-p104">必需。指向用户可以了解你的外接程序详细信息的页面 URL。**resid** 属性引用 [Resources](resources.md) 分区的 **Urls** 元素中的有效 ID。</span><span class="sxs-lookup"><span data-stu-id="77ef7-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="77ef7-131">**LearnMoreUrl** 当前无法在 Word、Excel 或 PowerPoint 客户端中呈现。</span><span class="sxs-lookup"><span data-stu-id="77ef7-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="77ef7-132">我们建议为所有客户端添加此 URL，以便 URL 在可用时呈现。</span><span class="sxs-lookup"><span data-stu-id="77ef7-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="77ef7-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="77ef7-133">See also</span></span>

<span data-ttu-id="77ef7-134">下面的代码示例使用 **GetStarted** 元素：</span><span class="sxs-lookup"><span data-stu-id="77ef7-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="77ef7-135">用于控制表和图表格式化的 Excel Web 外接程序</span><span class="sxs-lookup"><span data-stu-id="77ef7-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="77ef7-136">Word 外接程序 JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="77ef7-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="77ef7-137">在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表</span><span class="sxs-lookup"><span data-stu-id="77ef7-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
