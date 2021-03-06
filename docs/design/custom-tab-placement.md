---
title: 在功能区上定位自定义选项卡
description: 了解如何控制自定义选项卡在 Office 功能区上的显示位置以及默认情况下是否具有焦点。
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 6718a69191d1d84d96512c01b2544094ce276ab6
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505204"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a><span data-ttu-id="eb415-103">在功能区上定位自定义选项卡</span><span class="sxs-lookup"><span data-stu-id="eb415-103">Position a custom tab on the ribbon</span></span>

<span data-ttu-id="eb415-104">可以使用加载项清单中的标记指定希望外接程序的自定义选项卡在 Office 应用程序功能区上的显示位置。</span><span class="sxs-lookup"><span data-stu-id="eb415-104">You can specify where you want your add-in's custom tab to appear on the Office application's ribbon by using markup in the add-in's manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="eb415-105">本文假定您熟悉外接程序 [命令的基本概念一文](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="eb415-105">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="eb415-106">如果你最近没有这样做，请查看它。</span><span class="sxs-lookup"><span data-stu-id="eb415-106">Please review it if you have not done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="eb415-107">本文中介绍的外接程序功能及标记 *仅在 PowerPoint 网页中可用*。</span><span class="sxs-lookup"><span data-stu-id="eb415-107">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="eb415-108">本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。</span><span class="sxs-lookup"><span data-stu-id="eb415-108">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="eb415-109">请参阅 [下面的不受支持平台上的行为](#behavior-on-unsupported-platforms) 。</span><span class="sxs-lookup"><span data-stu-id="eb415-109">See [Behavior on unsupported platforms](#behavior-on-unsupported-platforms) below.</span></span>

<span data-ttu-id="eb415-110">通过标识希望自定义选项卡位于哪个内置 Office 选项卡旁边并指定自定义选项卡应位于内置选项卡的左侧还是右侧，来指定自定义选项卡的显示位置。在外接程序清单的[CustomTab](../reference/manifest/customtab.md)元素中添加[InsertBefore](../reference/manifest/customtab.md#insertbefore) () 或[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 元素，以创建这些规范。</span><span class="sxs-lookup"><span data-stu-id="eb415-110">Specify where you want a custom tab to appear by identifying which built-in Office tab you want it to be next to and specifying whether it should be on the left or right side of the built-in tab. Make these specifications by including either an [InsertBefore](../reference/manifest/customtab.md#insertbefore) (left) or an [InsertAfter](../reference/manifest/customtab.md#insertafter) (right) element in the [CustomTab](../reference/manifest/customtab.md) element of your add-in's manifest.</span></span> <span data-ttu-id="eb415-111"> (不能同时具有这两个元素。) </span><span class="sxs-lookup"><span data-stu-id="eb415-111">(You cannot have both elements.)</span></span>

<span data-ttu-id="eb415-112">在下面的示例中，自定义选项卡配置为显示在"审阅"*选项卡\*\*\*的正*\* 后。请注意，该元素 `<InsertAfter>` 的值是内置 Office 选项卡的 ID。</span><span class="sxs-lookup"><span data-stu-id="eb415-112">In the following example, the custom tab is configured to appear *just after* the **Review** tab. Note that the value of the `<InsertAfter>` element is the ID of the built-in Office tab.</span></span> 

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

<span data-ttu-id="eb415-113">请记住以下几点。</span><span class="sxs-lookup"><span data-stu-id="eb415-113">Keep the following points in mind.</span></span>

- <span data-ttu-id="eb415-114">And  `<InsertBefore>`  `<InsertAfter>` 元素是可选的。</span><span class="sxs-lookup"><span data-stu-id="eb415-114">The  `<InsertBefore>` and  `<InsertAfter>` elements are optional.</span></span> <span data-ttu-id="eb415-115">如果两者均不使用，则自定义选项卡将显示为功能区最右边的选项卡。</span><span class="sxs-lookup"><span data-stu-id="eb415-115">If you use neither, then your custom tab will appear as the rightmost tab on the ribbon.</span></span>
- <span data-ttu-id="eb415-116">和  `<InsertBefore>`  `<InsertAfter>` 元素相互排斥。</span><span class="sxs-lookup"><span data-stu-id="eb415-116">The  `<InsertBefore>` and  `<InsertAfter>` elements are mutually exclusive.</span></span> <span data-ttu-id="eb415-117">不能同时使用这两者。</span><span class="sxs-lookup"><span data-stu-id="eb415-117">You cannot use both.</span></span>
- <span data-ttu-id="eb415-118">如果用户安装了多个自定义选项卡配置为同一位置的外接程序（例如，在"审阅"选项卡之后，则最近安装的外接程序的选项卡将位于该位置）。</span><span class="sxs-lookup"><span data-stu-id="eb415-118">If the user installs more than one add-in whose custom tab is configured for the same place, say after the **Review** tab, then the tab for the most recently installed add-in will be located in that place.</span></span> <span data-ttu-id="eb415-119">以前安装的加载项的选项卡将移动到一处。</span><span class="sxs-lookup"><span data-stu-id="eb415-119">The tabs of the previously installed add-ins will be moved over one place.</span></span> <span data-ttu-id="eb415-120">例如，用户按该顺序安装外接程序 A、B 和 C，并且所有加载项均配置为在"审阅"选项卡后插入选项卡，然后选项卡将按以下顺序显示：Review、AddinCTab、AddinBTab、AddinATab。    </span><span class="sxs-lookup"><span data-stu-id="eb415-120">For example, the user installs add-ins A, B, and C in that order and all are configured to insert a tab after the **Review** tab, then the tabs will appear in this order: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.</span></span>
- <span data-ttu-id="eb415-121">用户可以在 Office 应用程序中自定义功能区。</span><span class="sxs-lookup"><span data-stu-id="eb415-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="eb415-122">例如，用户可以移动或隐藏加载项的选项卡。无法阻止此情况或检测到已发生此情况。</span><span class="sxs-lookup"><span data-stu-id="eb415-122">For example, a user can move or hide your add-in's tab. You cannot prevent this or detect that it has happened.</span></span>
- <span data-ttu-id="eb415-123">如果用户移动其中一个内置选项卡，则 Office 根据内置选项卡的默认位置解释 and `<InsertBefore>` `<InsertAfter>` *元素*。例如，如果用户将"审阅"选项卡移到功能区的右端，Office 会将上述示例中的标记解释为"将自定义选项卡放在"审阅"选项卡默认位置的右侧。" \*\*</span><span class="sxs-lookup"><span data-stu-id="eb415-123">If a user moves one of the built-in tabs, then Office interprets the `<InsertBefore>` and  `<InsertAfter>` elements in terms of *the default location of the built-in tab*. For example, if the user moves the **Review** tab to the right end of the ribbon, Office will interpret the markup in the example above as meaning "put the custom tab just to the right of *where the **Review** tab would be by default*."</span></span>

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a><span data-ttu-id="eb415-124">指定文档打开时哪个选项卡具有焦点</span><span class="sxs-lookup"><span data-stu-id="eb415-124">Specifying which tab has focus when the document opens</span></span>

<span data-ttu-id="eb415-125">Office 始终为紧接在"文件"选项卡右边的选项卡提供 **默认** 焦点。默认情况下，这是"主页 **"** 选项卡。如果将自定义选项卡配置为在"开始"选项卡之前，则打开文档时您的自定义选项卡 `<InsertBefore>TabHome</InsertBefore>` 将具有焦点。</span><span class="sxs-lookup"><span data-stu-id="eb415-125">Office always gives default focus to the tab that is immediately to the right of the **File** tab. By default this is the **Home** tab. If you configure your custom tab to be before the **Home** tab, with `<InsertBefore>TabHome</InsertBefore>`, then your custom tab will have focus when the document opens.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eb415-126">过分强调加载项的不便，并惹恼用户和管理员。</span><span class="sxs-lookup"><span data-stu-id="eb415-126">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span> <span data-ttu-id="eb415-127">除非外接程序是用户与文档交互的主要方式，否则不要将自定义选项卡定位到"主页"选项卡之前。</span><span class="sxs-lookup"><span data-stu-id="eb415-127">Do not position a custom tab before the **Home** tab unless your add-in is the primary way users will interact with the document.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="eb415-128">不受支持的平台上的行为</span><span class="sxs-lookup"><span data-stu-id="eb415-128">Behavior on unsupported platforms</span></span>

<span data-ttu-id="eb415-129">如果加载项安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则本文中描述的标记将被忽略，并且您的自定义选项卡将显示为功能区最右边的选项卡。</span><span class="sxs-lookup"><span data-stu-id="eb415-129">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and your custom tab will appear as the rightmost tab on the ribbon.</span></span> <span data-ttu-id="eb415-130">若要防止加载项安装在不支持标记的平台上，请添加对清单部分中的要求集 `<Requirements>` 的引用。</span><span class="sxs-lookup"><span data-stu-id="eb415-130">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="eb415-131">有关说明，请参阅 [清单中的设置 Requirements 元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="eb415-131">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="eb415-132">或者，可以将外接程序设计成在 **不支持 AddinCommands 1.3** 时具有备用体验，如 [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)代码中的"使用运行时检查"中所述。</span><span class="sxs-lookup"><span data-stu-id="eb415-132">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="eb415-133">例如，如果您的外接程序包含假定自定义选项卡位于需要它的说明，则您可能有一个备用版本，假定该选项卡最右侧。</span><span class="sxs-lookup"><span data-stu-id="eb415-133">For example, if your add-in contains instructions that assume the custom tab is where you want it, you could have an alternate version that assumes the tab is the rightmost.</span></span>
