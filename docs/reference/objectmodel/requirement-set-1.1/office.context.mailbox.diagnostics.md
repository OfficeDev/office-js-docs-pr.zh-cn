---
title: "\"Context.subname\": \"邮箱\"。诊断-要求集1。1"
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 0f0f23b28d32e1a4910082269e27138262262706
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127441"
---
# <a name="diagnostics"></a><span data-ttu-id="c0c82-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="c0c82-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="c0c82-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="c0c82-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="c0c82-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="c0c82-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0c82-105">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-105">Requirements</span></span>

|<span data-ttu-id="c0c82-106">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-106">Requirement</span></span>| <span data-ttu-id="c0c82-107">值</span><span class="sxs-lookup"><span data-stu-id="c0c82-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c82-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c82-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c82-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c82-109">1.0</span></span>|
|[<span data-ttu-id="c0c82-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0c82-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0c82-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0c82-111">ReadItem</span></span>|
|[<span data-ttu-id="c0c82-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c82-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c82-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c82-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="c0c82-114">Members</span><span class="sxs-lookup"><span data-stu-id="c0c82-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="c0c82-115">hostName: String</span><span class="sxs-lookup"><span data-stu-id="c0c82-115">hostName: String</span></span>

<span data-ttu-id="c0c82-116">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="c0c82-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="c0c82-117">可以是下列值之一的字符串：`Outlook`、`OutlookIOS` 或 `OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="c0c82-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c82-118">类型</span><span class="sxs-lookup"><span data-stu-id="c0c82-118">Type</span></span>

*   <span data-ttu-id="c0c82-119">String</span><span class="sxs-lookup"><span data-stu-id="c0c82-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0c82-120">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-120">Requirements</span></span>

|<span data-ttu-id="c0c82-121">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-121">Requirement</span></span>| <span data-ttu-id="c0c82-122">值</span><span class="sxs-lookup"><span data-stu-id="c0c82-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c82-123">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c82-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c82-124">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c82-124">1.0</span></span>|
|[<span data-ttu-id="c0c82-125">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0c82-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0c82-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0c82-126">ReadItem</span></span>|
|[<span data-ttu-id="c0c82-127">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c82-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c82-128">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c82-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="c0c82-129">Diagnostics.hostversion: String</span><span class="sxs-lookup"><span data-stu-id="c0c82-129">hostVersion: String</span></span>

<span data-ttu-id="c0c82-130">获取表示主机应用程序或 Exchange Server 的版本的字符串。</span><span class="sxs-lookup"><span data-stu-id="c0c82-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="c0c82-131">如果邮件外接程序在 Outlook 桌面客户端或 iOS 上运行, 则该`hostVersion`属性返回主机应用程序 (Outlook) 的版本。</span><span class="sxs-lookup"><span data-stu-id="c0c82-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="c0c82-132">在 Outlook 网页版中, 该属性返回的是 Exchange 服务器的版本。</span><span class="sxs-lookup"><span data-stu-id="c0c82-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="c0c82-133">例如，字符串 `15.0.468.0`。</span><span class="sxs-lookup"><span data-stu-id="c0c82-133">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c82-134">类型</span><span class="sxs-lookup"><span data-stu-id="c0c82-134">Type</span></span>

*   <span data-ttu-id="c0c82-135">String</span><span class="sxs-lookup"><span data-stu-id="c0c82-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0c82-136">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-136">Requirements</span></span>

|<span data-ttu-id="c0c82-137">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-137">Requirement</span></span>| <span data-ttu-id="c0c82-138">值</span><span class="sxs-lookup"><span data-stu-id="c0c82-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c82-139">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c82-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c82-140">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c82-140">1.0</span></span>|
|[<span data-ttu-id="c0c82-141">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0c82-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0c82-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0c82-142">ReadItem</span></span>|
|[<span data-ttu-id="c0c82-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c82-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c82-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c82-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="c0c82-145">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="c0c82-145">OWAView: String</span></span>

<span data-ttu-id="c0c82-146">获取表示 web 上的 Outlook 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="c0c82-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="c0c82-147">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="c0c82-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="c0c82-148">如果主机应用程序不是 web 上的 Outlook, 则访问此属性将导致`undefined`。</span><span class="sxs-lookup"><span data-stu-id="c0c82-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="c0c82-149">Web 上的 Outlook 具有三个视图, 分别对应于屏幕的宽度和窗口, 以及可以显示的列数:</span><span class="sxs-lookup"><span data-stu-id="c0c82-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="c0c82-150">`OneColumn` 在屏幕较窄时显示。</span><span class="sxs-lookup"><span data-stu-id="c0c82-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="c0c82-151">Web 上的 Outlook 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="c0c82-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="c0c82-152">`TwoColumns` 在屏幕较宽时显示。</span><span class="sxs-lookup"><span data-stu-id="c0c82-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="c0c82-153">Outlook 网页版在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="c0c82-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="c0c82-154">`ThreeColumns` 在屏幕为宽屏时显示。</span><span class="sxs-lookup"><span data-stu-id="c0c82-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="c0c82-155">例如, web 上的 Outlook 在桌面计算机上的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="c0c82-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c82-156">类型</span><span class="sxs-lookup"><span data-stu-id="c0c82-156">Type</span></span>

*   <span data-ttu-id="c0c82-157">String</span><span class="sxs-lookup"><span data-stu-id="c0c82-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0c82-158">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-158">Requirements</span></span>

|<span data-ttu-id="c0c82-159">要求</span><span class="sxs-lookup"><span data-stu-id="c0c82-159">Requirement</span></span>| <span data-ttu-id="c0c82-160">值</span><span class="sxs-lookup"><span data-stu-id="c0c82-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c82-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c82-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c82-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c82-162">1.0</span></span>|
|[<span data-ttu-id="c0c82-163">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0c82-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0c82-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0c82-164">ReadItem</span></span>|
|[<span data-ttu-id="c0c82-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c82-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c82-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c82-166">Compose or Read</span></span>|
