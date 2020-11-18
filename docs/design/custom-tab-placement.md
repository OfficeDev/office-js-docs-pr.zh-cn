---
title: 在功能区上定位自定义选项卡
description: 了解如何在默认情况下控制自定义选项卡在 Office 功能区上显示的位置以及它是否有焦点。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 2c1e2ae66805212e78868cf7c07a0e5c14cd4025
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088165"
---
# <a name="position-a-custom-tab-on-the-ribbon-preview"></a><span data-ttu-id="24b5a-103">将自定义选项卡放置在功能区上 (预览) </span><span class="sxs-lookup"><span data-stu-id="24b5a-103">Position a custom tab on the ribbon (preview)</span></span>

<span data-ttu-id="24b5a-104">您可以使用外接程序清单中的标记来指定您希望外接程序的自定义选项卡在 Office 应用程序的功能区上显示的位置。</span><span class="sxs-lookup"><span data-stu-id="24b5a-104">You can specify where you want your add-in's custom tab to appear on the Office application's ribbon by using markup in the add-in's manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="24b5a-105">本文假定您熟悉文章 [外接程序命令的基本概念](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="24b5a-105">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="24b5a-106">如果你最近未执行此操作，请查看它。</span><span class="sxs-lookup"><span data-stu-id="24b5a-106">Please review it if you have not done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="24b5a-107">本文中介绍的加载项功能和标记位于预览中， *仅适用于 web 上的 PowerPoint*。</span><span class="sxs-lookup"><span data-stu-id="24b5a-107">The add-in feature and markup described in this article is in preview and is *only available in PowerPoint on the web*.</span></span> <span data-ttu-id="24b5a-108">我们建议您仅在测试和开发环境中尝试标记。</span><span class="sxs-lookup"><span data-stu-id="24b5a-108">We recommend that you try out the markup in test and development environments only.</span></span> <span data-ttu-id="24b5a-109">请勿在生产环境中或在业务关键型文档中使用预览标记。</span><span class="sxs-lookup"><span data-stu-id="24b5a-109">Do not use preview markup in a production environment or within business-critical documents.</span></span>
> - <span data-ttu-id="24b5a-110">本文中所述的标记仅适用于支持要求集 **addincommand 1.3** 的平台。</span><span class="sxs-lookup"><span data-stu-id="24b5a-110">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="24b5a-111">请参阅下面 [有关不受支持的平台的行为](#behavior-on-unsupported-platforms) 。</span><span class="sxs-lookup"><span data-stu-id="24b5a-111">See [Behavior on unsupported platforms](#behavior-on-unsupported-platforms) below.</span></span>

<span data-ttu-id="24b5a-112">指定要显示自定义选项卡的位置，具体方法是确定您希望它在其旁边的内置 "Office" 选项卡，并指定它是在内置选项卡的左侧还是右侧。通过在外接程序清单的[CustomTab](../reference/manifest/customtab.md)元素中包括一个[InsertBefore](../reference/manifest/customtab.md#insertbefore) (left) 或[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 元素来设置这些规范。</span><span class="sxs-lookup"><span data-stu-id="24b5a-112">Specify where you want a custom tab to appear by identifying which built-in Office tab you want it to be next to and specifying whether it should be on the left or right side of the built-in tab. Make these specifications by including either an [InsertBefore](../reference/manifest/customtab.md#insertbefore) (left) or an [InsertAfter](../reference/manifest/customtab.md#insertafter) (right) element in the [CustomTab](../reference/manifest/customtab.md) element of your add-in's manifest.</span></span> <span data-ttu-id="24b5a-113"> (不能同时具有这两个元素 ) </span><span class="sxs-lookup"><span data-stu-id="24b5a-113">(You cannot have both elements.)</span></span>

<span data-ttu-id="24b5a-114">在以下示例中，将自定义选项卡配置为 *恰好* 显示在 " **审阅** " 选项卡的后面。请注意，该元素的值 `<InsertAfter>` 是内置 "Office" 选项卡的 ID。</span><span class="sxs-lookup"><span data-stu-id="24b5a-114">In the following example, the custom tab is configured to appear *just after* the **Review** tab. Note that the value of the `<InsertAfter>` element is the ID of the built-in Office tab.</span></span> 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

<span data-ttu-id="24b5a-115">请记住以下几点。</span><span class="sxs-lookup"><span data-stu-id="24b5a-115">Keep the following points in mind.</span></span>

- <span data-ttu-id="24b5a-116">`<InsertBefore>`和 `<InsertAfter>` 元素是可选的。</span><span class="sxs-lookup"><span data-stu-id="24b5a-116">The  `<InsertBefore>` and  `<InsertAfter>` elements are optional.</span></span> <span data-ttu-id="24b5a-117">如果不使用这两种方式，则自定义选项卡将显示为功能区上最右边的选项卡。</span><span class="sxs-lookup"><span data-stu-id="24b5a-117">If you use neither, then your custom tab will appear as the rightmost tab on the ribbon.</span></span>
- <span data-ttu-id="24b5a-118">`<InsertBefore>`和 `<InsertAfter>` 元素相互排斥。</span><span class="sxs-lookup"><span data-stu-id="24b5a-118">The  `<InsertBefore>` and  `<InsertAfter>` elements are mutually exclusive.</span></span> <span data-ttu-id="24b5a-119">您不能同时使用这两种。</span><span class="sxs-lookup"><span data-stu-id="24b5a-119">You cannot use both.</span></span>
- <span data-ttu-id="24b5a-120">如果用户安装了多个加载项，其自定义选项卡配置为相同位置，则在 " **审阅** " 选项卡之后，最近安装的加载项的选项卡将位于该位置。</span><span class="sxs-lookup"><span data-stu-id="24b5a-120">If the user installs more than one add-in whose custom tab is configured for the same place, say after the **Review** tab, then the tab for the most recently installed add-in will be located in that place.</span></span> <span data-ttu-id="24b5a-121">之前安装的外接程序的选项卡将移动到一个位置。</span><span class="sxs-lookup"><span data-stu-id="24b5a-121">The tabs of the previously installed add-ins will be moved over one place.</span></span> <span data-ttu-id="24b5a-122">例如，用户在该顺序中安装外接程序 A、B 和 C，并将所有配置为在 " **审阅** " 选项卡上将其配置为插入一个选项卡，然后选项卡将按如下顺序显示： " **审阅**"、" **AddinCTab**"、" **AddinBTab**"、" **AddinATab**"。</span><span class="sxs-lookup"><span data-stu-id="24b5a-122">For example, the user installs add-ins A, B, and C in that order and all are configured to insert a tab after the **Review** tab, then the tabs will appear in this order: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.</span></span>
- <span data-ttu-id="24b5a-123">用户可以在 Office 应用程序中自定义功能区。</span><span class="sxs-lookup"><span data-stu-id="24b5a-123">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="24b5a-124">例如，用户可以移动或隐藏外接程序的选项卡。您不能阻止此情况，也不能检测到此问题。</span><span class="sxs-lookup"><span data-stu-id="24b5a-124">For example, a user can move or hide your add-in's tab. You cannot prevent this or detect that it has happened.</span></span>
- <span data-ttu-id="24b5a-125">如果用户移动了其中一个内置选项卡，则 Office 将 `<InsertBefore>` `<InsertAfter>` 根据 *内置选项卡的默认位置* 来解释和元素。例如，如果用户将 "**审阅**" 选项卡移到功能区的右端，则 Office 会将上面示例中的标记解释为 "将自定义选项卡放在 ***审阅** 选项卡的默认位置* 的右侧"。</span><span class="sxs-lookup"><span data-stu-id="24b5a-125">If a user moves one of the built-in tabs, then Office interprets the `<InsertBefore>` and  `<InsertAfter>` elements in terms of *the default location of the built-in tab*. For example, if the user moves the **Review** tab to the right end of the ribbon, Office will interpret the markup in the example above as meaning "put the custom tab just to the right of *where the **Review** tab would be by default*."</span></span>

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a><span data-ttu-id="24b5a-126">指定在文档打开时哪个选项卡具有焦点</span><span class="sxs-lookup"><span data-stu-id="24b5a-126">Specifying which tab has focus when the document opens</span></span>

<span data-ttu-id="24b5a-127">Office 始终向紧靠 " **文件** " 选项卡右侧的选项卡提供默认焦点。默认情况下，这是 " **主页** " 选项卡。如果将自定义选项卡配置为在 " **主页** " 选项卡之前使用 `<InsertBefore>TabHome</InsertBefore>` ，则在打开文档时，自定义选项卡将获得焦点。</span><span class="sxs-lookup"><span data-stu-id="24b5a-127">Office always gives default focus to the tab that is immediately to the right of the **File** tab. By default this is the **Home** tab. If you configure your custom tab to be before the **Home** tab, with `<InsertBefore>TabHome</InsertBefore>`, then your custom tab will have focus when the document opens.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="24b5a-128">为您的外接程序 inconveniences 和 annoys 用户和管理员提供了更多的突出。</span><span class="sxs-lookup"><span data-stu-id="24b5a-128">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span> <span data-ttu-id="24b5a-129">不要将自定义选项卡放置在 " **主页** " 选项卡之前，除非您的外接程序是用户将与文档进行交互的主要方式。</span><span class="sxs-lookup"><span data-stu-id="24b5a-129">Do not position a custom tab before the **Home** tab unless your add-in is the primary way users will interact with the document.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="24b5a-130">不受支持的平台上的行为</span><span class="sxs-lookup"><span data-stu-id="24b5a-130">Behavior on unsupported platforms</span></span>

<span data-ttu-id="24b5a-131">如果您的外接程序安装在不支持 [要求集 addincommand 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则会忽略本文中所述的标记，并且您的自定义选项卡将显示为功能区上最右边的选项卡。</span><span class="sxs-lookup"><span data-stu-id="24b5a-131">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and your custom tab will appear as the rightmost tab on the ribbon.</span></span> <span data-ttu-id="24b5a-132">若要防止外接程序安装在不支持标记的平台上，请在清单的部分中添加对要求集的引用 `<Requirements>` 。</span><span class="sxs-lookup"><span data-stu-id="24b5a-132">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="24b5a-133">有关说明，请参阅 [在清单中设置需求元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="24b5a-133">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="24b5a-134">此外，还可以设计外接程序，使其在 **addincommand 1.3** 不受支持时具有备用体验，如 [JavaScript 代码中的使用运行时检查](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)中所述。</span><span class="sxs-lookup"><span data-stu-id="24b5a-134">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="24b5a-135">例如，如果您的外接程序包含假设自定义选项卡位于您所需的位置的说明，则可以使用该选项卡的备选版本，该选项卡位于最右边。</span><span class="sxs-lookup"><span data-stu-id="24b5a-135">For example, if your add-in contains instructions that assume the custom tab is where you want it, you could have an alternate version that assumes the tab is the rightmost.</span></span>
