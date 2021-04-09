---
title: 清单文件中运行时
description: Runtime 元素将外接程序配置为将共享 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652242"
---
# <a name="runtime-element"></a><span data-ttu-id="26ac7-103">运行时元素</span><span class="sxs-lookup"><span data-stu-id="26ac7-103">Runtime element</span></span>

<span data-ttu-id="26ac7-104">将外接程序配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="26ac7-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="26ac7-105">元素的 [`<Runtimes>`](runtimes.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="26ac7-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="26ac7-106">**外接程序类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="26ac7-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="26ac7-107">语法</span><span class="sxs-lookup"><span data-stu-id="26ac7-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="26ac7-108">包含于</span><span class="sxs-lookup"><span data-stu-id="26ac7-108">Contained in</span></span>

- [<span data-ttu-id="26ac7-109">运行时</span><span class="sxs-lookup"><span data-stu-id="26ac7-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="26ac7-110">属性</span><span class="sxs-lookup"><span data-stu-id="26ac7-110">Attributes</span></span>

|  <span data-ttu-id="26ac7-111">属性</span><span class="sxs-lookup"><span data-stu-id="26ac7-111">Attribute</span></span>  |  <span data-ttu-id="26ac7-112">必需</span><span class="sxs-lookup"><span data-stu-id="26ac7-112">Required</span></span>  |  <span data-ttu-id="26ac7-113">说明</span><span class="sxs-lookup"><span data-stu-id="26ac7-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="26ac7-114">**resid**</span><span class="sxs-lookup"><span data-stu-id="26ac7-114">**resid**</span></span>  |  <span data-ttu-id="26ac7-115">是</span><span class="sxs-lookup"><span data-stu-id="26ac7-115">Yes</span></span>  | <span data-ttu-id="26ac7-116">指定外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="26ac7-116">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="26ac7-117">`resid`不能超过 32 个字符，并且必须与 元素中的 `id` `Url` 元素的 属性 `Resources` 匹配。</span><span class="sxs-lookup"><span data-stu-id="26ac7-117">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="26ac7-118">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="26ac7-118">**lifetime**</span></span>  |  <span data-ttu-id="26ac7-119">否</span><span class="sxs-lookup"><span data-stu-id="26ac7-119">No</span></span>  | <span data-ttu-id="26ac7-120">的默认值是 `lifetime` `short` ，不需要指定。</span><span class="sxs-lookup"><span data-stu-id="26ac7-120">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="26ac7-121">Outlook 外接程序仅使用 `short` 值。</span><span class="sxs-lookup"><span data-stu-id="26ac7-121">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="26ac7-122">如果要在 Excel 加载项中使用共享运行时，请显式将值设置为 `long` 。</span><span class="sxs-lookup"><span data-stu-id="26ac7-122">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="26ac7-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="26ac7-123">See also</span></span>

- [<span data-ttu-id="26ac7-124">运行时</span><span class="sxs-lookup"><span data-stu-id="26ac7-124">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="26ac7-125">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="26ac7-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="26ac7-126">配置 Outlook 外接程序进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="26ac7-126">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
