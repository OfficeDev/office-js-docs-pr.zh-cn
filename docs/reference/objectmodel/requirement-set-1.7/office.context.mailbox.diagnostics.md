---
title: "\"Context.subname\": \"邮箱\"。诊断-要求集1。7"
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 3f17f21bd4fcf48496ba7f387a57752256d4a038
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695964"
---
# <a name="diagnostics"></a><span data-ttu-id="02c88-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="02c88-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="02c88-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="02c88-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="02c88-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="02c88-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02c88-105">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-105">Requirements</span></span>

|<span data-ttu-id="02c88-106">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-106">Requirement</span></span>| <span data-ttu-id="02c88-107">值</span><span class="sxs-lookup"><span data-stu-id="02c88-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="02c88-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02c88-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02c88-109">1.0</span><span class="sxs-lookup"><span data-stu-id="02c88-109">1.0</span></span>|
|[<span data-ttu-id="02c88-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02c88-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02c88-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02c88-111">ReadItem</span></span>|
|[<span data-ttu-id="02c88-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02c88-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02c88-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02c88-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="02c88-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="02c88-114">Members and methods</span></span>

| <span data-ttu-id="02c88-115">成员</span><span class="sxs-lookup"><span data-stu-id="02c88-115">Member</span></span> | <span data-ttu-id="02c88-116">类型</span><span class="sxs-lookup"><span data-stu-id="02c88-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="02c88-117">主机名</span><span class="sxs-lookup"><span data-stu-id="02c88-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="02c88-118">Member</span><span class="sxs-lookup"><span data-stu-id="02c88-118">Member</span></span> |
| [<span data-ttu-id="02c88-119">Diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="02c88-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="02c88-120">Member</span><span class="sxs-lookup"><span data-stu-id="02c88-120">Member</span></span> |
| [<span data-ttu-id="02c88-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="02c88-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="02c88-122">Member</span><span class="sxs-lookup"><span data-stu-id="02c88-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="02c88-123">Members</span><span class="sxs-lookup"><span data-stu-id="02c88-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="02c88-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="02c88-124">hostName: String</span></span>

<span data-ttu-id="02c88-125">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="02c88-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="02c88-126">可以是下列值之一的字符串：`Outlook`、`OutlookWebApp`、`OutlookIOS` 或 `OutlookAndroid`。</span><span class="sxs-lookup"><span data-stu-id="02c88-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="02c88-127">对`Outlook`桌面客户端 (即 Windows 和 Mac) 上的 Outlook 返回值。</span><span class="sxs-lookup"><span data-stu-id="02c88-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="02c88-128">类型</span><span class="sxs-lookup"><span data-stu-id="02c88-128">Type</span></span>

*   <span data-ttu-id="02c88-129">String</span><span class="sxs-lookup"><span data-stu-id="02c88-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02c88-130">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-130">Requirements</span></span>

|<span data-ttu-id="02c88-131">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-131">Requirement</span></span>| <span data-ttu-id="02c88-132">值</span><span class="sxs-lookup"><span data-stu-id="02c88-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="02c88-133">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02c88-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02c88-134">1.0</span><span class="sxs-lookup"><span data-stu-id="02c88-134">1.0</span></span>|
|[<span data-ttu-id="02c88-135">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02c88-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02c88-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02c88-136">ReadItem</span></span>|
|[<span data-ttu-id="02c88-137">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02c88-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02c88-138">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02c88-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="02c88-139">Diagnostics.hostversion: String</span><span class="sxs-lookup"><span data-stu-id="02c88-139">hostVersion: String</span></span>

<span data-ttu-id="02c88-140">获取表示主机应用程序或 Exchange 服务器的版本的字符串 (例如, "15.0.468.0")。</span><span class="sxs-lookup"><span data-stu-id="02c88-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="02c88-141">如果邮件外接程序在 Outlook 桌面客户端或 iOS 上运行, 则该`hostVersion`属性返回主机应用程序 (Outlook) 的版本。</span><span class="sxs-lookup"><span data-stu-id="02c88-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="02c88-142">在 Outlook 网页版中, 该属性返回的是 Exchange 服务器的版本。</span><span class="sxs-lookup"><span data-stu-id="02c88-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="02c88-143">类型</span><span class="sxs-lookup"><span data-stu-id="02c88-143">Type</span></span>

*   <span data-ttu-id="02c88-144">String</span><span class="sxs-lookup"><span data-stu-id="02c88-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02c88-145">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-145">Requirements</span></span>

|<span data-ttu-id="02c88-146">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-146">Requirement</span></span>| <span data-ttu-id="02c88-147">值</span><span class="sxs-lookup"><span data-stu-id="02c88-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="02c88-148">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02c88-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02c88-149">1.0</span><span class="sxs-lookup"><span data-stu-id="02c88-149">1.0</span></span>|
|[<span data-ttu-id="02c88-150">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02c88-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02c88-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02c88-151">ReadItem</span></span>|
|[<span data-ttu-id="02c88-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02c88-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02c88-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02c88-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="02c88-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="02c88-154">OWAView: String</span></span>

<span data-ttu-id="02c88-155">获取表示 web 上的 Outlook 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="02c88-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="02c88-156">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="02c88-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="02c88-157">如果主机应用程序不是 web 上的 Outlook, 则访问此属性将导致`undefined`。</span><span class="sxs-lookup"><span data-stu-id="02c88-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="02c88-158">Web 上的 Outlook 具有三个视图, 分别对应于屏幕的宽度和窗口, 以及可以显示的列数:</span><span class="sxs-lookup"><span data-stu-id="02c88-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="02c88-159">`OneColumn` 在屏幕较窄时显示。</span><span class="sxs-lookup"><span data-stu-id="02c88-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="02c88-160">Web 上的 Outlook 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="02c88-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="02c88-161">`TwoColumns` 在屏幕较宽时显示。</span><span class="sxs-lookup"><span data-stu-id="02c88-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="02c88-162">Outlook 网页版在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="02c88-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="02c88-163">`ThreeColumns` 在屏幕为宽屏时显示。</span><span class="sxs-lookup"><span data-stu-id="02c88-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="02c88-164">例如, web 上的 Outlook 在桌面计算机上的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="02c88-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="02c88-165">类型</span><span class="sxs-lookup"><span data-stu-id="02c88-165">Type</span></span>

*   <span data-ttu-id="02c88-166">String</span><span class="sxs-lookup"><span data-stu-id="02c88-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02c88-167">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-167">Requirements</span></span>

|<span data-ttu-id="02c88-168">要求</span><span class="sxs-lookup"><span data-stu-id="02c88-168">Requirement</span></span>| <span data-ttu-id="02c88-169">值</span><span class="sxs-lookup"><span data-stu-id="02c88-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="02c88-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02c88-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02c88-171">1.0</span><span class="sxs-lookup"><span data-stu-id="02c88-171">1.0</span></span>|
|[<span data-ttu-id="02c88-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02c88-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02c88-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02c88-173">ReadItem</span></span>|
|[<span data-ttu-id="02c88-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02c88-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02c88-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02c88-175">Compose or Read</span></span>|
