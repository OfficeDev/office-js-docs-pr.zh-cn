---
title: Office 外接程序的导航模式
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: f0482f7742c6fab97fe563d61d21135c072f8f8f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449131"
---
# <a name="navigation-patterns"></a><span data-ttu-id="39e68-102">导航模式</span><span class="sxs-lookup"><span data-stu-id="39e68-102">Navigation patterns</span></span>

<span data-ttu-id="39e68-103">可以通过特定命令类型和指定的屏幕区域访问外接程序的主要功能。</span><span class="sxs-lookup"><span data-stu-id="39e68-103">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="39e68-104">导航直观明了，可提供上下文并允许用户在外接程序中轻松移动，这些非常重要。</span><span class="sxs-lookup"><span data-stu-id="39e68-104">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="39e68-105">最佳做法</span><span class="sxs-lookup"><span data-stu-id="39e68-105">Best practices</span></span>

| <span data-ttu-id="39e68-106">允许事项</span><span class="sxs-lookup"><span data-stu-id="39e68-106">Do</span></span>    | <span data-ttu-id="39e68-107">禁止事项</span><span class="sxs-lookup"><span data-stu-id="39e68-107">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="39e68-108">确保为用户提供清晰的可视化导航选项。</span><span class="sxs-lookup"><span data-stu-id="39e68-108">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="39e68-109">不要使用非标准 UI，使导航过程变得复杂。</span><span class="sxs-lookup"><span data-stu-id="39e68-109">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="39e68-110">使用以下组件（如适用）允许用户在加载程序中导航。</span><span class="sxs-lookup"><span data-stu-id="39e68-110">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="39e68-111">不要让用户难以知悉其当前在外接程序中所处的位置或上下文</span><span class="sxs-lookup"><span data-stu-id="39e68-111">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="39e68-112">命令栏</span><span class="sxs-lookup"><span data-stu-id="39e68-112">Command Bar</span></span>

<span data-ttu-id="39e68-113">命令栏是一个图面，其中包含在其驻留的窗口、面板或父区域内容上运行的命令。</span><span class="sxs-lookup"><span data-stu-id="39e68-113">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="39e68-114">可选功能包括汉堡菜单访问点、搜索和侧命令。</span><span class="sxs-lookup"><span data-stu-id="39e68-114">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![命令 - 桌面任务窗格规范](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="39e68-116">选项卡栏</span><span class="sxs-lookup"><span data-stu-id="39e68-116">Tab Bar</span></span>

<span data-ttu-id="39e68-117">显示使用具有垂直堆叠文本和图标的按钮进行导航。</span><span class="sxs-lookup"><span data-stu-id="39e68-117">Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="39e68-118">使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。</span><span class="sxs-lookup"><span data-stu-id="39e68-118">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![选项卡栏 - 桌面任务窗格规范](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="39e68-120">“返回”按钮</span><span class="sxs-lookup"><span data-stu-id="39e68-120">Back Button</span></span>

<span data-ttu-id="39e68-121">“返回”按钮使用户能够恢复向下钻取导航操作。</span><span class="sxs-lookup"><span data-stu-id="39e68-121">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="39e68-122">此模式有助于确保用户遵循一系列有序的步骤。</span><span class="sxs-lookup"><span data-stu-id="39e68-122">This pattern helps ensure users follow an ordered series of steps.</span></span>  

![“返回”按钮 - 桌面任务窗格规范](../images/add-in-back-button.png)
