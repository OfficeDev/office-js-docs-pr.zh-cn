---
title: 清单文件中的 GetStarted 元素
description: 提供在 Word、Excel、PowerPoint 和 OneNote 中安装加载项时出现的标注使用的信息。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0ad6196dc45e4ea06c2b43ac5da66a560ab0b899
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771412"
---
# <a name="getstarted-element"></a><span data-ttu-id="a56e3-103">GetStarted 元素</span><span class="sxs-lookup"><span data-stu-id="a56e3-103">GetStarted element</span></span>

<span data-ttu-id="a56e3-104">提供在 Word、Excel、PowerPoint 和 OneNote 中安装加载项时出现的标注使用的信息。</span><span class="sxs-lookup"><span data-stu-id="a56e3-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="a56e3-105">**GetStarted** 元素是 [DesktopFormFactor](desktopformfactor.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="a56e3-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="a56e3-106">子元素</span><span class="sxs-lookup"><span data-stu-id="a56e3-106">Child elements</span></span>

| <span data-ttu-id="a56e3-107">元素</span><span class="sxs-lookup"><span data-stu-id="a56e3-107">Element</span></span>                       | <span data-ttu-id="a56e3-108">必需</span><span class="sxs-lookup"><span data-stu-id="a56e3-108">Required</span></span> | <span data-ttu-id="a56e3-109">说明</span><span class="sxs-lookup"><span data-stu-id="a56e3-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="a56e3-110">标题</span><span class="sxs-lookup"><span data-stu-id="a56e3-110">Title</span></span>](#title)               | <span data-ttu-id="a56e3-111">是</span><span class="sxs-lookup"><span data-stu-id="a56e3-111">Yes</span></span>      | <span data-ttu-id="a56e3-112">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="a56e3-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="a56e3-113">说明</span><span class="sxs-lookup"><span data-stu-id="a56e3-113">Description</span></span>](#description)   | <span data-ttu-id="a56e3-114">是</span><span class="sxs-lookup"><span data-stu-id="a56e3-114">Yes</span></span>      | <span data-ttu-id="a56e3-115">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="a56e3-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="a56e3-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="a56e3-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="a56e3-117">是</span><span class="sxs-lookup"><span data-stu-id="a56e3-117">Yes</span></span>       | <span data-ttu-id="a56e3-118">指向详细说明外接程序的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="a56e3-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="a56e3-119">标题</span><span class="sxs-lookup"><span data-stu-id="a56e3-119">Title</span></span> 

<span data-ttu-id="a56e3-120">必需。</span><span class="sxs-lookup"><span data-stu-id="a56e3-120">Required.</span></span> <span data-ttu-id="a56e3-121">用于标注顶部的标题。</span><span class="sxs-lookup"><span data-stu-id="a56e3-121">The title used for the top of the callout.</span></span> <span data-ttu-id="a56e3-122">resid 属性引用"资源"部分 **ShortStrings** 元素中的 [](resources.md)有效 ID，并且不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="a56e3-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="a56e3-123">说明</span><span class="sxs-lookup"><span data-stu-id="a56e3-123">Description</span></span>

<span data-ttu-id="a56e3-124">必需。</span><span class="sxs-lookup"><span data-stu-id="a56e3-124">Required.</span></span> <span data-ttu-id="a56e3-125">标注的说明/正文内容。</span><span class="sxs-lookup"><span data-stu-id="a56e3-125">The description / body content for the callout.</span></span> <span data-ttu-id="a56e3-126">resid 属性引用"资源"部分 **LongStrings** 元素中的 [](resources.md)有效 ID，并且不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="a56e3-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="a56e3-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="a56e3-127">LearnMoreUrl</span></span>

<span data-ttu-id="a56e3-128">必需。</span><span class="sxs-lookup"><span data-stu-id="a56e3-128">Required.</span></span> <span data-ttu-id="a56e3-129">指向用户可以了解你的外接程序详细信息的页面 URL。</span><span class="sxs-lookup"><span data-stu-id="a56e3-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="a56e3-130">resid 属性引用 Resources 节 **的 Urls** 元素 [](resources.md)中的有效 ID，并且不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="a56e3-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="a56e3-131">**LearnMoreUrl** 当前无法在 Word、Excel 或 PowerPoint 客户端中呈现。</span><span class="sxs-lookup"><span data-stu-id="a56e3-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="a56e3-132">我们建议为所有客户端添加此 URL，以便 URL 在可用时呈现。</span><span class="sxs-lookup"><span data-stu-id="a56e3-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="a56e3-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a56e3-133">See also</span></span>

<span data-ttu-id="a56e3-134">下面的代码示例使用 **GetStarted** 元素：</span><span class="sxs-lookup"><span data-stu-id="a56e3-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="a56e3-135">用于控制表和图表格式化的 Excel Web 外接程序</span><span class="sxs-lookup"><span data-stu-id="a56e3-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="a56e3-136">Word 外接程序 JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="a56e3-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="a56e3-137">在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表</span><span class="sxs-lookup"><span data-stu-id="a56e3-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
