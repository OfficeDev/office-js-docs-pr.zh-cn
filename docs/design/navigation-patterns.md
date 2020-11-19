---
title: Office 外接程序的导航模式
description: 了解使用命令栏、选项卡栏和后退按钮的最佳实践，以设计 Office 外接程序的导航。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132030"
---
# <a name="navigation-patterns"></a><span data-ttu-id="267ff-103">导航模式</span><span class="sxs-lookup"><span data-stu-id="267ff-103">Navigation patterns</span></span>

<span data-ttu-id="267ff-104">可以通过特定命令类型和指定的屏幕区域访问外接程序的主要功能。</span><span class="sxs-lookup"><span data-stu-id="267ff-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="267ff-105">导航直观明了，可提供上下文并允许用户在外接程序中轻松移动，这些非常重要。</span><span class="sxs-lookup"><span data-stu-id="267ff-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="267ff-106">最佳做法</span><span class="sxs-lookup"><span data-stu-id="267ff-106">Best practices</span></span>

| <span data-ttu-id="267ff-107">允许事项</span><span class="sxs-lookup"><span data-stu-id="267ff-107">Do</span></span>    | <span data-ttu-id="267ff-108">禁止事项</span><span class="sxs-lookup"><span data-stu-id="267ff-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="267ff-109">确保为用户提供清晰的可视化导航选项。</span><span class="sxs-lookup"><span data-stu-id="267ff-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="267ff-110">不要使用非标准 UI，使导航过程变得复杂。</span><span class="sxs-lookup"><span data-stu-id="267ff-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="267ff-111">使用以下组件（如适用）允许用户在加载程序中导航。</span><span class="sxs-lookup"><span data-stu-id="267ff-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="267ff-112">不要让用户难以知悉其当前在外接程序中所处的位置或上下文</span><span class="sxs-lookup"><span data-stu-id="267ff-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>

## <a name="command-bar"></a><span data-ttu-id="267ff-113">命令栏</span><span class="sxs-lookup"><span data-stu-id="267ff-113">Command Bar</span></span>

<span data-ttu-id="267ff-114">命令栏是任务窗格中的一个图面，其中驻留了在其驻留的窗口、面板或父区域的内容上运行的命令。</span><span class="sxs-lookup"><span data-stu-id="267ff-114">The CommandBar is a surface within the task pane that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="267ff-115">可选功能包括汉堡菜单访问点、搜索和侧命令。</span><span class="sxs-lookup"><span data-stu-id="267ff-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![图示显示在 "Office 桌面应用程序" 任务窗格中的命令栏。](../images/add-in-command-bar.png)

## <a name="tab-bar"></a><span data-ttu-id="267ff-118">选项卡栏</span><span class="sxs-lookup"><span data-stu-id="267ff-118">Tab Bar</span></span>

<span data-ttu-id="267ff-119">选项卡栏显示了使用垂直堆叠文本和图标的按钮的导航。</span><span class="sxs-lookup"><span data-stu-id="267ff-119">The tab bar shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="267ff-120">使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。</span><span class="sxs-lookup"><span data-stu-id="267ff-120">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![图示显示在 "Office 桌面应用程序" 任务窗格中的选项卡栏。](../images/add-in-tab-bar.png)

## <a name="back-button"></a><span data-ttu-id="267ff-123">“返回”按钮</span><span class="sxs-lookup"><span data-stu-id="267ff-123">Back Button</span></span>

<span data-ttu-id="267ff-124">"后退" 按钮允许用户从深化导航操作中恢复。</span><span class="sxs-lookup"><span data-stu-id="267ff-124">The back button allows users to recover from a drill-down navigational action.</span></span> <span data-ttu-id="267ff-125">此模式有助于确保用户遵循一系列有序的步骤。</span><span class="sxs-lookup"><span data-stu-id="267ff-125">This pattern helps ensure users follow an ordered series of steps.</span></span>

![显示 Office 桌面应用程序任务窗格中的 "后退" 按钮的图示。](../images/add-in-back-button.png)
