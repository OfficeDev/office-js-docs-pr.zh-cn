---
title: Office.context.mailbox.diagnostics - 要求集 1.4
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 3764bcf3629fb5a242c7c39df16f103e696613cf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433836"
---
# <a name="diagnostics"></a><span data-ttu-id="32372-102">诊断</span><span class="sxs-lookup"><span data-stu-id="32372-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="32372-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="32372-103">Office.context.mailbox.diagnostics</span></span>

<span data-ttu-id="32372-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="32372-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="32372-105">要求</span><span class="sxs-lookup"><span data-stu-id="32372-105">Requirements</span></span>

|<span data-ttu-id="32372-106">要求</span><span class="sxs-lookup"><span data-stu-id="32372-106">Requirement</span></span>| <span data-ttu-id="32372-107">值</span><span class="sxs-lookup"><span data-stu-id="32372-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="32372-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="32372-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32372-109">1.0</span><span class="sxs-lookup"><span data-stu-id="32372-109">1.0</span></span>|
|[<span data-ttu-id="32372-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="32372-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32372-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32372-111">ReadItem</span></span>|
|[<span data-ttu-id="32372-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="32372-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32372-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="32372-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="32372-114">成员</span><span class="sxs-lookup"><span data-stu-id="32372-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="32372-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="32372-115">hostName :String</span></span>

<span data-ttu-id="32372-116">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="32372-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="32372-117">可以是下列值之一的字符串：`Outlook`、`OutlookIOS` 或 `OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="32372-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="32372-118">类型:</span><span class="sxs-lookup"><span data-stu-id="32372-118">Type:</span></span>

*   <span data-ttu-id="32372-119">String</span><span class="sxs-lookup"><span data-stu-id="32372-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32372-120">要求</span><span class="sxs-lookup"><span data-stu-id="32372-120">Requirements</span></span>

|<span data-ttu-id="32372-121">要求</span><span class="sxs-lookup"><span data-stu-id="32372-121">Requirement</span></span>| <span data-ttu-id="32372-122">值</span><span class="sxs-lookup"><span data-stu-id="32372-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="32372-123">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="32372-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32372-124">1.0</span><span class="sxs-lookup"><span data-stu-id="32372-124">1.0</span></span>|
|[<span data-ttu-id="32372-125">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="32372-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32372-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32372-126">ReadItem</span></span>|
|[<span data-ttu-id="32372-127">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="32372-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32372-128">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="32372-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="32372-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="32372-129">hostVersion :String</span></span>

<span data-ttu-id="32372-130">获取表示主机应用程序或 Exchange Server 的版本的字符串。</span><span class="sxs-lookup"><span data-stu-id="32372-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="32372-p101">如果邮件外接程序正在 Outlook 桌面客户端或 Outlook for iOS 上运行，则 `hostVersion` 属性返回主机应用程序版本 Outlook。在 Outlook Web App 中，属性返回 Exchange Server 的版本。其中的一个示例是字符串 `15.0.468.0`。</span><span class="sxs-lookup"><span data-stu-id="32372-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="32372-134">类型:</span><span class="sxs-lookup"><span data-stu-id="32372-134">Type:</span></span>

*   <span data-ttu-id="32372-135">String</span><span class="sxs-lookup"><span data-stu-id="32372-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32372-136">要求</span><span class="sxs-lookup"><span data-stu-id="32372-136">Requirements</span></span>

|<span data-ttu-id="32372-137">要求</span><span class="sxs-lookup"><span data-stu-id="32372-137">Requirement</span></span>| <span data-ttu-id="32372-138">值</span><span class="sxs-lookup"><span data-stu-id="32372-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="32372-139">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="32372-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32372-140">1.0</span><span class="sxs-lookup"><span data-stu-id="32372-140">1.0</span></span>|
|[<span data-ttu-id="32372-141">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="32372-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32372-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32372-142">ReadItem</span></span>|
|[<span data-ttu-id="32372-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="32372-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32372-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="32372-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="32372-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="32372-145">OWAView :String</span></span>

<span data-ttu-id="32372-146">获取表示 Outlook Web App 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="32372-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="32372-147">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="32372-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="32372-148">如果主机应用程序不是 Outlook Web App，则访问此属性将导致返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="32372-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="32372-149">Outlook Web App 具有三种视图，这些视图分别与屏幕和窗口的宽度以及可以显示的列数相对应：</span><span class="sxs-lookup"><span data-stu-id="32372-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="32372-p102">`OneColumn` 在屏幕较窄时显示。Outlook Web App 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="32372-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="32372-p103">`TwoColumns` 在屏幕较宽时显示。Outlook Web App 在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="32372-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="32372-p104">`ThreeColumns` 在屏幕为宽屏时显示。例如，Outlook Web App 在台式机的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="32372-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="32372-156">类型：</span><span class="sxs-lookup"><span data-stu-id="32372-156">Type:</span></span>

*   <span data-ttu-id="32372-157">String</span><span class="sxs-lookup"><span data-stu-id="32372-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32372-158">要求</span><span class="sxs-lookup"><span data-stu-id="32372-158">Requirements</span></span>

|<span data-ttu-id="32372-159">要求</span><span class="sxs-lookup"><span data-stu-id="32372-159">Requirement</span></span>| <span data-ttu-id="32372-160">值</span><span class="sxs-lookup"><span data-stu-id="32372-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="32372-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="32372-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32372-162">1.0</span><span class="sxs-lookup"><span data-stu-id="32372-162">1.0</span></span>|
|[<span data-ttu-id="32372-163">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="32372-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32372-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32372-164">ReadItem</span></span>|
|[<span data-ttu-id="32372-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="32372-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32372-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="32372-166">Compose or read</span></span>|