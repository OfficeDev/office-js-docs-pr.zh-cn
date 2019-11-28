---
title: "\"Context.subname\"： \"邮箱\"。诊断-预览要求集"
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 492e292737417854adfaf98feb2b67788933d874
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629200"
---
# <a name="diagnostics"></a><span data-ttu-id="b366f-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="b366f-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="b366f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="b366f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="b366f-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="b366f-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b366f-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="b366f-105">Requirements</span></span>

|<span data-ttu-id="b366f-106">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-106">Requirement</span></span>| <span data-ttu-id="b366f-107">值</span><span class="sxs-lookup"><span data-stu-id="b366f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b366f-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b366f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b366f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-109">1.0</span></span>|
|[<span data-ttu-id="b366f-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b366f-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b366f-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-111">ReadItem</span></span>|
|[<span data-ttu-id="b366f-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b366f-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b366f-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b366f-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b366f-114">属性</span><span class="sxs-lookup"><span data-stu-id="b366f-114">Properties</span></span>

| <span data-ttu-id="b366f-115">属性</span><span class="sxs-lookup"><span data-stu-id="b366f-115">Property</span></span> | <span data-ttu-id="b366f-116">最低</span><span class="sxs-lookup"><span data-stu-id="b366f-116">Minimum</span></span><br><span data-ttu-id="b366f-117">权限级别</span><span class="sxs-lookup"><span data-stu-id="b366f-117">permission level</span></span> | <span data-ttu-id="b366f-118">型号</span><span class="sxs-lookup"><span data-stu-id="b366f-118">Modes</span></span> | <span data-ttu-id="b366f-119">返回类型</span><span class="sxs-lookup"><span data-stu-id="b366f-119">Return type</span></span> | <span data-ttu-id="b366f-120">最低</span><span class="sxs-lookup"><span data-stu-id="b366f-120">Minimum</span></span><br><span data-ttu-id="b366f-121">要求集</span><span class="sxs-lookup"><span data-stu-id="b366f-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="b366f-122">主机名</span><span class="sxs-lookup"><span data-stu-id="b366f-122">hostName</span></span>](#hostname-string) | <span data-ttu-id="b366f-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-123">ReadItem</span></span> | <span data-ttu-id="b366f-124">撰写</span><span class="sxs-lookup"><span data-stu-id="b366f-124">Compose</span></span><br><span data-ttu-id="b366f-125">读取</span><span class="sxs-lookup"><span data-stu-id="b366f-125">Read</span></span> | <span data-ttu-id="b366f-126">String</span><span class="sxs-lookup"><span data-stu-id="b366f-126">String</span></span> | <span data-ttu-id="b366f-127">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-127">1.0</span></span> |
| [<span data-ttu-id="b366f-128">Diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="b366f-128">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="b366f-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-129">ReadItem</span></span> | <span data-ttu-id="b366f-130">撰写</span><span class="sxs-lookup"><span data-stu-id="b366f-130">Compose</span></span><br><span data-ttu-id="b366f-131">读取</span><span class="sxs-lookup"><span data-stu-id="b366f-131">Read</span></span> | <span data-ttu-id="b366f-132">String</span><span class="sxs-lookup"><span data-stu-id="b366f-132">String</span></span> | <span data-ttu-id="b366f-133">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-133">1.0</span></span> |
| [<span data-ttu-id="b366f-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="b366f-134">OWAView</span></span>](#owaview-string) | <span data-ttu-id="b366f-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-135">ReadItem</span></span> | <span data-ttu-id="b366f-136">撰写</span><span class="sxs-lookup"><span data-stu-id="b366f-136">Compose</span></span><br><span data-ttu-id="b366f-137">读取</span><span class="sxs-lookup"><span data-stu-id="b366f-137">Read</span></span> | <span data-ttu-id="b366f-138">String</span><span class="sxs-lookup"><span data-stu-id="b366f-138">String</span></span> | <span data-ttu-id="b366f-139">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-139">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="b366f-140">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="b366f-140">Property details</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="b366f-141">hostName： String</span><span class="sxs-lookup"><span data-stu-id="b366f-141">hostName: String</span></span>

<span data-ttu-id="b366f-142">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="b366f-142">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="b366f-143">可以是下列值之一的字符串：`Outlook`、`OutlookWebApp`、`OutlookIOS` 或 `OutlookAndroid`。</span><span class="sxs-lookup"><span data-stu-id="b366f-143">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="b366f-144">对`Outlook`桌面客户端（即 Windows 和 Mac）上的 Outlook 返回值。</span><span class="sxs-lookup"><span data-stu-id="b366f-144">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="b366f-145">类型</span><span class="sxs-lookup"><span data-stu-id="b366f-145">Type</span></span>

*   <span data-ttu-id="b366f-146">String</span><span class="sxs-lookup"><span data-stu-id="b366f-146">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b366f-147">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-147">Requirements</span></span>

|<span data-ttu-id="b366f-148">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-148">Requirement</span></span>| <span data-ttu-id="b366f-149">值</span><span class="sxs-lookup"><span data-stu-id="b366f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="b366f-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b366f-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b366f-151">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-151">1.0</span></span>|
|[<span data-ttu-id="b366f-152">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b366f-152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b366f-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-153">ReadItem</span></span>|
|[<span data-ttu-id="b366f-154">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b366f-154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b366f-155">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b366f-155">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="b366f-156">Diagnostics.hostversion： String</span><span class="sxs-lookup"><span data-stu-id="b366f-156">hostVersion: String</span></span>

<span data-ttu-id="b366f-157">获取表示主机应用程序或 Exchange 服务器的版本的字符串（例如，"15.0.468.0"）。</span><span class="sxs-lookup"><span data-stu-id="b366f-157">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="b366f-158">如果邮件外接程序在 Outlook 桌面或移动客户端上运行，则该`hostVersion`属性将返回主机应用程序（Outlook）的版本。</span><span class="sxs-lookup"><span data-stu-id="b366f-158">If the mail add-in is running on an Outlook desktop or mobile client, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="b366f-159">在 Outlook 网页版中，该属性返回的是 Exchange 服务器的版本。</span><span class="sxs-lookup"><span data-stu-id="b366f-159">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="b366f-160">类型</span><span class="sxs-lookup"><span data-stu-id="b366f-160">Type</span></span>

*   <span data-ttu-id="b366f-161">String</span><span class="sxs-lookup"><span data-stu-id="b366f-161">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b366f-162">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-162">Requirements</span></span>

|<span data-ttu-id="b366f-163">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-163">Requirement</span></span>| <span data-ttu-id="b366f-164">值</span><span class="sxs-lookup"><span data-stu-id="b366f-164">Value</span></span>|
|---|---|
|[<span data-ttu-id="b366f-165">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b366f-165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b366f-166">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-166">1.0</span></span>|
|[<span data-ttu-id="b366f-167">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b366f-167">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b366f-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-168">ReadItem</span></span>|
|[<span data-ttu-id="b366f-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b366f-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b366f-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b366f-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="b366f-171">OWAView： String</span><span class="sxs-lookup"><span data-stu-id="b366f-171">OWAView: String</span></span>

<span data-ttu-id="b366f-172">获取表示 web 上的 Outlook 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="b366f-172">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="b366f-173">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="b366f-173">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="b366f-174">如果主机应用程序不是 web 上的 Outlook，则访问此属性将导致`undefined`。</span><span class="sxs-lookup"><span data-stu-id="b366f-174">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="b366f-175">Web 上的 Outlook 具有三个视图，分别对应于屏幕的宽度和窗口，以及可以显示的列数：</span><span class="sxs-lookup"><span data-stu-id="b366f-175">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="b366f-176">`OneColumn` 在屏幕较窄时显示。</span><span class="sxs-lookup"><span data-stu-id="b366f-176">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="b366f-177">Web 上的 Outlook 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="b366f-177">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="b366f-178">`TwoColumns` 在屏幕较宽时显示。</span><span class="sxs-lookup"><span data-stu-id="b366f-178">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="b366f-179">Outlook 网页版在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="b366f-179">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="b366f-180">`ThreeColumns` 在屏幕为宽屏时显示。</span><span class="sxs-lookup"><span data-stu-id="b366f-180">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="b366f-181">例如，web 上的 Outlook 在桌面计算机上的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="b366f-181">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="b366f-182">类型</span><span class="sxs-lookup"><span data-stu-id="b366f-182">Type</span></span>

*   <span data-ttu-id="b366f-183">String</span><span class="sxs-lookup"><span data-stu-id="b366f-183">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b366f-184">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-184">Requirements</span></span>

|<span data-ttu-id="b366f-185">要求</span><span class="sxs-lookup"><span data-stu-id="b366f-185">Requirement</span></span>| <span data-ttu-id="b366f-186">值</span><span class="sxs-lookup"><span data-stu-id="b366f-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="b366f-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b366f-187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b366f-188">1.0</span><span class="sxs-lookup"><span data-stu-id="b366f-188">1.0</span></span>|
|[<span data-ttu-id="b366f-189">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b366f-189">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b366f-190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b366f-190">ReadItem</span></span>|
|[<span data-ttu-id="b366f-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b366f-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b366f-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b366f-192">Compose or Read</span></span>|
