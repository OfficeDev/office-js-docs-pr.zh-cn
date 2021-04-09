---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652230"
---
# <a name="runtimes-element"></a><span data-ttu-id="b5886-103">Runtimes 元素</span><span class="sxs-lookup"><span data-stu-id="b5886-103">Runtimes element</span></span>

<span data-ttu-id="b5886-104">指定外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="b5886-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="b5886-105">元素的 [`<Host>`](host.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="b5886-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="b5886-106">When running in Office on Windows， your add-in uses the Internet Explorer 11 browser.</span><span class="sxs-lookup"><span data-stu-id="b5886-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="b5886-107">**外接程序类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="b5886-107">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="b5886-108">语法</span><span class="sxs-lookup"><span data-stu-id="b5886-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b5886-109">包含于</span><span class="sxs-lookup"><span data-stu-id="b5886-109">Contained in</span></span>

[<span data-ttu-id="b5886-110">Host</span><span class="sxs-lookup"><span data-stu-id="b5886-110">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="b5886-111">子元素</span><span class="sxs-lookup"><span data-stu-id="b5886-111">Child elements</span></span>

|  <span data-ttu-id="b5886-112">元素</span><span class="sxs-lookup"><span data-stu-id="b5886-112">Element</span></span> |  <span data-ttu-id="b5886-113">必需</span><span class="sxs-lookup"><span data-stu-id="b5886-113">Required</span></span>  |  <span data-ttu-id="b5886-114">说明</span><span class="sxs-lookup"><span data-stu-id="b5886-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="b5886-115">运行时</span><span class="sxs-lookup"><span data-stu-id="b5886-115">Runtime</span></span>](runtime.md) | <span data-ttu-id="b5886-116">是</span><span class="sxs-lookup"><span data-stu-id="b5886-116">Yes</span></span> |  <span data-ttu-id="b5886-117">加载项的运行时。</span><span class="sxs-lookup"><span data-stu-id="b5886-117">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b5886-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b5886-118">See also</span></span>

- [<span data-ttu-id="b5886-119">运行时</span><span class="sxs-lookup"><span data-stu-id="b5886-119">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="b5886-120">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="b5886-120">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="b5886-121">配置 Outlook 外接程序进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="b5886-121">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
