---
title: Office.context.mailbox.diagnostics - 要求集 1.6
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: a3b64e2c24c28fc93807ce3ab6261111e44e28b4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432709"
---
# <a name="diagnostics"></a><span data-ttu-id="dad3a-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="dad3a-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="dad3a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="dad3a-103">Office.context.mailbox.diagnostics</span></span>

<span data-ttu-id="dad3a-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="dad3a-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dad3a-105">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-105">Requirements</span></span>

|<span data-ttu-id="dad3a-106">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-106">Requirement</span></span>| <span data-ttu-id="dad3a-107">值</span><span class="sxs-lookup"><span data-stu-id="dad3a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dad3a-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dad3a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dad3a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dad3a-109">1.0</span></span>|
|[<span data-ttu-id="dad3a-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="dad3a-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dad3a-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dad3a-111">ReadItem</span></span>|
|[<span data-ttu-id="dad3a-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dad3a-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dad3a-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dad3a-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dad3a-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="dad3a-114">Members and methods</span></span>

| <span data-ttu-id="dad3a-115">成员</span><span class="sxs-lookup"><span data-stu-id="dad3a-115">Member</span></span> | <span data-ttu-id="dad3a-116">类型</span><span class="sxs-lookup"><span data-stu-id="dad3a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dad3a-117">hostName</span><span class="sxs-lookup"><span data-stu-id="dad3a-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="dad3a-118">成员</span><span class="sxs-lookup"><span data-stu-id="dad3a-118">Member</span></span> |
| [<span data-ttu-id="dad3a-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="dad3a-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="dad3a-120">成员</span><span class="sxs-lookup"><span data-stu-id="dad3a-120">Member</span></span> |
| [<span data-ttu-id="dad3a-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="dad3a-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="dad3a-122">成员</span><span class="sxs-lookup"><span data-stu-id="dad3a-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="dad3a-123">成员</span><span class="sxs-lookup"><span data-stu-id="dad3a-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="dad3a-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="dad3a-124">hostName :String</span></span>

<span data-ttu-id="dad3a-125">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="dad3a-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="dad3a-126">可以是下列值之一的字符串：`Outlook`、`Mac Outlook`、`OutlookIOS` 或 `OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="dad3a-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="dad3a-127">类型:</span><span class="sxs-lookup"><span data-stu-id="dad3a-127">Type:</span></span>

*   <span data-ttu-id="dad3a-128">String</span><span class="sxs-lookup"><span data-stu-id="dad3a-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dad3a-129">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-129">Requirements</span></span>

|<span data-ttu-id="dad3a-130">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-130">Requirement</span></span>| <span data-ttu-id="dad3a-131">值</span><span class="sxs-lookup"><span data-stu-id="dad3a-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="dad3a-132">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dad3a-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dad3a-133">1.0</span><span class="sxs-lookup"><span data-stu-id="dad3a-133">1.0</span></span>|
|[<span data-ttu-id="dad3a-134">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="dad3a-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dad3a-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dad3a-135">ReadItem</span></span>|
|[<span data-ttu-id="dad3a-136">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dad3a-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dad3a-137">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dad3a-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="dad3a-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="dad3a-138">hostVersion :String</span></span>

<span data-ttu-id="dad3a-139">获取表示主机应用程序或 Exchange Server 的版本的字符串。</span><span class="sxs-lookup"><span data-stu-id="dad3a-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="dad3a-p101">如果邮件外接程序正在 Outlook 桌面客户端或 Outlook for iOS 上运行，则 `hostVersion` 属性返回主机应用程序版本 Outlook。在 Outlook Web App 中，属性返回 Exchange Server 的版本。其中的一个示例是字符串 `15.0.468.0`。</span><span class="sxs-lookup"><span data-stu-id="dad3a-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="dad3a-143">类型:</span><span class="sxs-lookup"><span data-stu-id="dad3a-143">Type:</span></span>

*   <span data-ttu-id="dad3a-144">String</span><span class="sxs-lookup"><span data-stu-id="dad3a-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dad3a-145">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-145">Requirements</span></span>

|<span data-ttu-id="dad3a-146">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-146">Requirement</span></span>| <span data-ttu-id="dad3a-147">值</span><span class="sxs-lookup"><span data-stu-id="dad3a-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="dad3a-148">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dad3a-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dad3a-149">1.0</span><span class="sxs-lookup"><span data-stu-id="dad3a-149">1.0</span></span>|
|[<span data-ttu-id="dad3a-150">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="dad3a-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dad3a-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dad3a-151">ReadItem</span></span>|
|[<span data-ttu-id="dad3a-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dad3a-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dad3a-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dad3a-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="dad3a-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="dad3a-154">OWAView :String</span></span>

<span data-ttu-id="dad3a-155">获取表示 Outlook Web App 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="dad3a-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="dad3a-156">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="dad3a-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="dad3a-157">如果主机应用程序不是 Outlook Web App，则访问此属性将导致返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="dad3a-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="dad3a-158">Outlook Web App 具有三种视图，这些视图分别与屏幕和窗口的宽度以及可以显示的列数相对应：</span><span class="sxs-lookup"><span data-stu-id="dad3a-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="dad3a-p102">`OneColumn` 在屏幕较窄时显示。Outlook Web App 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="dad3a-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="dad3a-p103">`TwoColumns` 在屏幕较宽时显示。Outlook Web App 在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="dad3a-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="dad3a-p104">`ThreeColumns` 在屏幕为宽屏时显示。例如，Outlook Web App 在台式机的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="dad3a-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="dad3a-165">类型：</span><span class="sxs-lookup"><span data-stu-id="dad3a-165">Type:</span></span>

*   <span data-ttu-id="dad3a-166">String</span><span class="sxs-lookup"><span data-stu-id="dad3a-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dad3a-167">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-167">Requirements</span></span>

|<span data-ttu-id="dad3a-168">要求</span><span class="sxs-lookup"><span data-stu-id="dad3a-168">Requirement</span></span>| <span data-ttu-id="dad3a-169">值</span><span class="sxs-lookup"><span data-stu-id="dad3a-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="dad3a-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dad3a-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dad3a-171">1.0</span><span class="sxs-lookup"><span data-stu-id="dad3a-171">1.0</span></span>|
|[<span data-ttu-id="dad3a-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="dad3a-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dad3a-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dad3a-173">ReadItem</span></span>|
|[<span data-ttu-id="dad3a-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dad3a-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dad3a-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dad3a-175">Compose or read</span></span>|