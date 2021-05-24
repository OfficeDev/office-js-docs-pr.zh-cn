---
title: 清单文件中运行时
description: Runtime 元素将外接程序配置为将共享 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590910"
---
# <a name="runtime-element"></a><span data-ttu-id="b1fd8-103">运行时元素</span><span class="sxs-lookup"><span data-stu-id="b1fd8-103">Runtime element</span></span>

<span data-ttu-id="b1fd8-104">将外接程序配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="b1fd8-105">元素的 [`<Runtimes>`](runtimes.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="b1fd8-106">**外接程序类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="b1fd8-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="b1fd8-107">语法</span><span class="sxs-lookup"><span data-stu-id="b1fd8-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b1fd8-108">包含于</span><span class="sxs-lookup"><span data-stu-id="b1fd8-108">Contained in</span></span>

- [<span data-ttu-id="b1fd8-109">运行时</span><span class="sxs-lookup"><span data-stu-id="b1fd8-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="b1fd8-110">子元素</span><span class="sxs-lookup"><span data-stu-id="b1fd8-110">Child elements</span></span>

|  <span data-ttu-id="b1fd8-111">元素</span><span class="sxs-lookup"><span data-stu-id="b1fd8-111">Element</span></span> |  <span data-ttu-id="b1fd8-112">必需</span><span class="sxs-lookup"><span data-stu-id="b1fd8-112">Required</span></span>  |  <span data-ttu-id="b1fd8-113">说明</span><span class="sxs-lookup"><span data-stu-id="b1fd8-113">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="b1fd8-114">Override</span><span class="sxs-lookup"><span data-stu-id="b1fd8-114">Override</span></span>](override.md) | <span data-ttu-id="b1fd8-115">否</span><span class="sxs-lookup"><span data-stu-id="b1fd8-115">No</span></span> | <span data-ttu-id="b1fd8-116">**Outlook**：指定 Desktop 为 [LaunchEvent](../../reference/manifest/extensionpoint.md#launchevent)扩展点处理程序Outlook JavaScript 文件的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span> <span data-ttu-id="b1fd8-117">**重要** 提示：目前只能定义一 `<Override>` 个元素，并且必须为 类型 `javascript` 。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="b1fd8-118">属性</span><span class="sxs-lookup"><span data-stu-id="b1fd8-118">Attributes</span></span>

|  <span data-ttu-id="b1fd8-119">属性</span><span class="sxs-lookup"><span data-stu-id="b1fd8-119">Attribute</span></span>  |  <span data-ttu-id="b1fd8-120">必需</span><span class="sxs-lookup"><span data-stu-id="b1fd8-120">Required</span></span>  |  <span data-ttu-id="b1fd8-121">说明</span><span class="sxs-lookup"><span data-stu-id="b1fd8-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b1fd8-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="b1fd8-122">**resid**</span></span>  |  <span data-ttu-id="b1fd8-123">是</span><span class="sxs-lookup"><span data-stu-id="b1fd8-123">Yes</span></span>  | <span data-ttu-id="b1fd8-124">指定外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="b1fd8-125">`resid`不能超过 32 个字符，并且必须与 元素中的 `id` `Url` 元素的 属性 `Resources` 匹配。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="b1fd8-126">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="b1fd8-126">**lifetime**</span></span>  |  <span data-ttu-id="b1fd8-127">否</span><span class="sxs-lookup"><span data-stu-id="b1fd8-127">No</span></span>  | <span data-ttu-id="b1fd8-128">的默认值是 `lifetime` `short` ，不需要指定。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="b1fd8-129">Outlook加载项只能使用 `short` 值。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="b1fd8-130">如果要在加载项中Excel运行时，请显式将值设置为 `long` 。</span><span class="sxs-lookup"><span data-stu-id="b1fd8-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b1fd8-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b1fd8-131">See also</span></span>

- [<span data-ttu-id="b1fd8-132">运行时</span><span class="sxs-lookup"><span data-stu-id="b1fd8-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="b1fd8-133">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="b1fd8-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="b1fd8-134">配置Outlook加载项进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="b1fd8-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
