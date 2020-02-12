---
title: Office 加载项中的对话框
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 6728e9032ba00c2e2ebcaa339f72700bc4dacca5
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950381"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="dd6c0-102">Office 加载项中的对话框</span><span class="sxs-lookup"><span data-stu-id="dd6c0-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="dd6c0-p101">对话框是浮动在活动的 Office 应用程序窗口之上的界面。你可以使用对话框为无法直接在任务窗格中打开的任务（例如登录页）提供额外的屏幕空间，或请求确认用户进行的操作，或显示如果局限在任务窗格中可能过小的视频。</span><span class="sxs-lookup"><span data-stu-id="dd6c0-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="dd6c0-105">*图 1：对话框典型布局*</span><span class="sxs-lookup"><span data-stu-id="dd6c0-105">*Figure 1. Typical layout for a dialog box*</span></span>

![显示对话框典型布局的示例图像](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="dd6c0-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="dd6c0-107">Best practices</span></span>

|<span data-ttu-id="dd6c0-108">**允许事项**</span><span class="sxs-lookup"><span data-stu-id="dd6c0-108">**Do**</span></span>|<span data-ttu-id="dd6c0-109">**禁止事项**</span><span class="sxs-lookup"><span data-stu-id="dd6c0-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="dd6c0-110">包括包含外接程序名称以及当前任务的描述性标题。</span><span class="sxs-lookup"><span data-stu-id="dd6c0-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="dd6c0-111">请勿在标题中追加公司名称。</span><span class="sxs-lookup"><span data-stu-id="dd6c0-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="dd6c0-112">除非方案需要，否则请勿打开对话框。</span><span class="sxs-lookup"><span data-stu-id="dd6c0-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="dd6c0-113">实现</span><span class="sxs-lookup"><span data-stu-id="dd6c0-113">Implementation</span></span>

<span data-ttu-id="dd6c0-114">有关实现对话框的示例，请参阅 GitHub 上的 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="dd6c0-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="dd6c0-115">另请参阅 </span><span class="sxs-lookup"><span data-stu-id="dd6c0-115">See also</span></span>

- [<span data-ttu-id="dd6c0-116">Dialog 对象</span><span class="sxs-lookup"><span data-stu-id="dd6c0-116">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="dd6c0-117">适用于 Office 加载项的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="dd6c0-117">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
