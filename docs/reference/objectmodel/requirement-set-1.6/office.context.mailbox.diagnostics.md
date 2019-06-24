---
title: "\"Context.subname\": \"邮箱\"。诊断-要求集1。6"
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 04fd4af8e35b2a538e93a64254250d40c3334dc6
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127371"
---
# <a name="diagnostics"></a><span data-ttu-id="60e56-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="60e56-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="60e56-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="60e56-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="60e56-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="60e56-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="60e56-105">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-105">Requirements</span></span>

|<span data-ttu-id="60e56-106">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-106">Requirement</span></span>| <span data-ttu-id="60e56-107">值</span><span class="sxs-lookup"><span data-stu-id="60e56-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="60e56-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60e56-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60e56-109">1.0</span><span class="sxs-lookup"><span data-stu-id="60e56-109">1.0</span></span>|
|[<span data-ttu-id="60e56-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60e56-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60e56-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60e56-111">ReadItem</span></span>|
|[<span data-ttu-id="60e56-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60e56-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60e56-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60e56-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="60e56-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="60e56-114">Members and methods</span></span>

| <span data-ttu-id="60e56-115">成员</span><span class="sxs-lookup"><span data-stu-id="60e56-115">Member</span></span> | <span data-ttu-id="60e56-116">类型</span><span class="sxs-lookup"><span data-stu-id="60e56-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="60e56-117">主机名</span><span class="sxs-lookup"><span data-stu-id="60e56-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="60e56-118">Member</span><span class="sxs-lookup"><span data-stu-id="60e56-118">Member</span></span> |
| [<span data-ttu-id="60e56-119">Diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="60e56-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="60e56-120">Member</span><span class="sxs-lookup"><span data-stu-id="60e56-120">Member</span></span> |
| [<span data-ttu-id="60e56-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="60e56-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="60e56-122">Member</span><span class="sxs-lookup"><span data-stu-id="60e56-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="60e56-123">Members</span><span class="sxs-lookup"><span data-stu-id="60e56-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="60e56-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="60e56-124">hostName: String</span></span>

<span data-ttu-id="60e56-125">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="60e56-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="60e56-126">可以是下列值之一的字符串：`Outlook`、`Mac Outlook`、`OutlookIOS` 或 `OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="60e56-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="60e56-127">类型</span><span class="sxs-lookup"><span data-stu-id="60e56-127">Type</span></span>

*   <span data-ttu-id="60e56-128">String</span><span class="sxs-lookup"><span data-stu-id="60e56-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60e56-129">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-129">Requirements</span></span>

|<span data-ttu-id="60e56-130">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-130">Requirement</span></span>| <span data-ttu-id="60e56-131">值</span><span class="sxs-lookup"><span data-stu-id="60e56-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="60e56-132">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60e56-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60e56-133">1.0</span><span class="sxs-lookup"><span data-stu-id="60e56-133">1.0</span></span>|
|[<span data-ttu-id="60e56-134">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60e56-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60e56-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60e56-135">ReadItem</span></span>|
|[<span data-ttu-id="60e56-136">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60e56-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60e56-137">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60e56-137">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="60e56-138">Diagnostics.hostversion: String</span><span class="sxs-lookup"><span data-stu-id="60e56-138">hostVersion: String</span></span>

<span data-ttu-id="60e56-139">获取表示主机应用程序或 Exchange Server 的版本的字符串。</span><span class="sxs-lookup"><span data-stu-id="60e56-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="60e56-140">如果邮件外接程序在 Outlook 桌面客户端或 iOS 上运行, 则该`hostVersion`属性返回主机应用程序 (Outlook) 的版本。</span><span class="sxs-lookup"><span data-stu-id="60e56-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="60e56-141">在 Outlook 网页版中, 该属性返回的是 Exchange 服务器的版本。</span><span class="sxs-lookup"><span data-stu-id="60e56-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="60e56-142">例如，字符串 `15.0.468.0`。</span><span class="sxs-lookup"><span data-stu-id="60e56-142">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="60e56-143">类型</span><span class="sxs-lookup"><span data-stu-id="60e56-143">Type</span></span>

*   <span data-ttu-id="60e56-144">String</span><span class="sxs-lookup"><span data-stu-id="60e56-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60e56-145">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-145">Requirements</span></span>

|<span data-ttu-id="60e56-146">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-146">Requirement</span></span>| <span data-ttu-id="60e56-147">值</span><span class="sxs-lookup"><span data-stu-id="60e56-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="60e56-148">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60e56-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60e56-149">1.0</span><span class="sxs-lookup"><span data-stu-id="60e56-149">1.0</span></span>|
|[<span data-ttu-id="60e56-150">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60e56-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60e56-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60e56-151">ReadItem</span></span>|
|[<span data-ttu-id="60e56-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60e56-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60e56-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60e56-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="60e56-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="60e56-154">OWAView: String</span></span>

<span data-ttu-id="60e56-155">获取表示 web 上的 Outlook 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="60e56-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="60e56-156">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="60e56-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="60e56-157">如果主机应用程序不是 web 上的 Outlook, 则访问此属性将导致`undefined`。</span><span class="sxs-lookup"><span data-stu-id="60e56-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="60e56-158">Web 上的 Outlook 具有三个视图, 分别对应于屏幕的宽度和窗口, 以及可以显示的列数:</span><span class="sxs-lookup"><span data-stu-id="60e56-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="60e56-159">`OneColumn` 在屏幕较窄时显示。</span><span class="sxs-lookup"><span data-stu-id="60e56-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="60e56-160">Web 上的 Outlook 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="60e56-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="60e56-161">`TwoColumns` 在屏幕较宽时显示。</span><span class="sxs-lookup"><span data-stu-id="60e56-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="60e56-162">Outlook 网页版在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="60e56-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="60e56-163">`ThreeColumns` 在屏幕为宽屏时显示。</span><span class="sxs-lookup"><span data-stu-id="60e56-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="60e56-164">例如, web 上的 Outlook 在桌面计算机上的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="60e56-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="60e56-165">类型</span><span class="sxs-lookup"><span data-stu-id="60e56-165">Type</span></span>

*   <span data-ttu-id="60e56-166">String</span><span class="sxs-lookup"><span data-stu-id="60e56-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60e56-167">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-167">Requirements</span></span>

|<span data-ttu-id="60e56-168">要求</span><span class="sxs-lookup"><span data-stu-id="60e56-168">Requirement</span></span>| <span data-ttu-id="60e56-169">值</span><span class="sxs-lookup"><span data-stu-id="60e56-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="60e56-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60e56-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60e56-171">1.0</span><span class="sxs-lookup"><span data-stu-id="60e56-171">1.0</span></span>|
|[<span data-ttu-id="60e56-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60e56-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60e56-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60e56-173">ReadItem</span></span>|
|[<span data-ttu-id="60e56-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60e56-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60e56-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60e56-175">Compose or Read</span></span>|
