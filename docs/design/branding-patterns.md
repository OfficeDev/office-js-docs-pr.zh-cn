---
title: Office 外接程序的品牌模式设计准则
description: ''
ms.date: 06/26/2018
ms.openlocfilehash: a94e723b222dfe1b004d8b558da59804faf51e69
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433698"
---
# <a name="branding-patterns"></a><span data-ttu-id="a5656-102">品牌模式</span><span class="sxs-lookup"><span data-stu-id="a5656-102">Branding patterns</span></span>

<span data-ttu-id="a5656-103">这些模式为外接程序用户提供了品牌可视化和上下文。</span><span class="sxs-lookup"><span data-stu-id="a5656-103">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="a5656-104">最佳做法</span><span class="sxs-lookup"><span data-stu-id="a5656-104">Best practices</span></span>

|<span data-ttu-id="a5656-105">允许事项</span><span class="sxs-lookup"><span data-stu-id="a5656-105">Do</span></span> |<span data-ttu-id="a5656-106">禁止事项</span><span class="sxs-lookup"><span data-stu-id="a5656-106">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="a5656-107">使用熟悉的 UI 组件并应用品牌个性色，如版式和颜色。</span><span class="sxs-lookup"><span data-stu-id="a5656-107">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="a5656-108">请勿使用与已创建的 Office UI 相冲突的新 UI 组价。</span><span class="sxs-lookup"><span data-stu-id="a5656-108">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="a5656-109">将外接程序品牌置于 UI 底部的品牌栏页脚中。</span><span class="sxs-lookup"><span data-stu-id="a5656-109">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="a5656-110">请勿在 UI 顶部的相邻品牌栏中重复任务窗格名称。</span><span class="sxs-lookup"><span data-stu-id="a5656-110">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="a5656-111">谨慎使用品牌元素。</span><span class="sxs-lookup"><span data-stu-id="a5656-111">Use brand elements sparingly.</span></span> <span data-ttu-id="a5656-112">将你的解决方案应用到 Office 中，使两者相得益彰。</span><span class="sxs-lookup"><span data-stu-id="a5656-112">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="a5656-113">不要再 Office UI 中插入过多的品牌元素，对用户造成干扰和迷惑。</span><span class="sxs-lookup"><span data-stu-id="a5656-113">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="a5656-114">使你的解决方案可识别，并利用一致的可视化元素将屏幕连接在一起。</span><span class="sxs-lookup"><span data-stu-id="a5656-114">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="a5656-115">不要让不可识别和不一致的可视化元素隐藏你的解决方案。</span><span class="sxs-lookup"><span data-stu-id="a5656-115">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="a5656-116">与父服务或业务建立连接，确保客户了解并信任你的解决方案。</span><span class="sxs-lookup"><span data-stu-id="a5656-116">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="a5656-117">如果可以通过有用且可理解的关系来建立信任和创造价值，则不要再给客户灌输新的品牌概念。</span><span class="sxs-lookup"><span data-stu-id="a5656-117">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="a5656-118">根据需要应用以下模式和组件，让用户充分利用外接程序的所有实用工具。</span><span class="sxs-lookup"><span data-stu-id="a5656-118">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="a5656-119">品牌栏</span><span class="sxs-lookup"><span data-stu-id="a5656-119">Brand Bar</span></span>

<span data-ttu-id="a5656-120">品牌栏是页脚中的一个区域，其中包含品牌名称和徽标。</span><span class="sxs-lookup"><span data-stu-id="a5656-120">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="a5656-121">此外，它还可以用作品牌网站和可选访问位置的链接。</span><span class="sxs-lookup"><span data-stu-id="a5656-121">It also serves as a link to your brand's website and an optional access location.</span></span>

![品牌栏 - 桌面任务窗格规范](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="a5656-123">初始屏幕</span><span class="sxs-lookup"><span data-stu-id="a5656-123">Splash Screen</span></span>

<span data-ttu-id="a5656-124">使用此屏幕在外接程序正在加载或转换 UI 状态时显示你的品牌。</span><span class="sxs-lookup"><span data-stu-id="a5656-124">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![品牌初始屏幕 - 桌面任务窗格规范](../images/add-in-splash-screen.png)