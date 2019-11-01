---
title: "\"Context\"-\"邮箱-预览要求集\""
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: ff649029713984b32e817bbeaf7c59a48cc5b023
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902107"
---
# <a name="mailbox"></a><span data-ttu-id="4f7fb-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="4f7fb-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4f7fb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4f7fb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4f7fb-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f7fb-105">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-105">Requirements</span></span>

|<span data-ttu-id="4f7fb-106">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-106">Requirement</span></span>| <span data-ttu-id="4f7fb-107">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-109">1.0</span></span>|
|[<span data-ttu-id="4f7fb-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-111">受限</span><span class="sxs-lookup"><span data-stu-id="4f7fb-111">Restricted</span></span>|
|[<span data-ttu-id="4f7fb-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4f7fb-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-114">Members and methods</span></span>

| <span data-ttu-id="4f7fb-115">成员</span><span class="sxs-lookup"><span data-stu-id="4f7fb-115">Member</span></span> | <span data-ttu-id="4f7fb-116">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4f7fb-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="4f7fb-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="4f7fb-118">成员</span><span class="sxs-lookup"><span data-stu-id="4f7fb-118">Member</span></span> |
| [<span data-ttu-id="4f7fb-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="4f7fb-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="4f7fb-120">成员</span><span class="sxs-lookup"><span data-stu-id="4f7fb-120">Member</span></span> |
| [<span data-ttu-id="4f7fb-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="4f7fb-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="4f7fb-122">成员</span><span class="sxs-lookup"><span data-stu-id="4f7fb-122">Member</span></span> |
| [<span data-ttu-id="4f7fb-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4f7fb-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4f7fb-124">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-124">Method</span></span> |
| [<span data-ttu-id="4f7fb-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="4f7fb-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="4f7fb-126">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-126">Method</span></span> |
| [<span data-ttu-id="4f7fb-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4f7fb-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="4f7fb-128">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-128">Method</span></span> |
| [<span data-ttu-id="4f7fb-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="4f7fb-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="4f7fb-130">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-130">Method</span></span> |
| [<span data-ttu-id="4f7fb-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="4f7fb-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="4f7fb-132">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-132">Method</span></span> |
| [<span data-ttu-id="4f7fb-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4f7fb-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="4f7fb-134">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-134">Method</span></span> |
| [<span data-ttu-id="4f7fb-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="4f7fb-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="4f7fb-136">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-136">Method</span></span> |
| [<span data-ttu-id="4f7fb-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4f7fb-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="4f7fb-138">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-138">Method</span></span> |
| [<span data-ttu-id="4f7fb-139">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="4f7fb-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="4f7fb-140">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-140">Method</span></span> |
| [<span data-ttu-id="4f7fb-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f7fb-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="4f7fb-142">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-142">Method</span></span> |
| [<span data-ttu-id="4f7fb-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f7fb-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="4f7fb-144">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-144">Method</span></span> |
| [<span data-ttu-id="4f7fb-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f7fb-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="4f7fb-146">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-146">Method</span></span> |
| [<span data-ttu-id="4f7fb-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="4f7fb-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="4f7fb-148">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-148">Method</span></span> |
| [<span data-ttu-id="4f7fb-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4f7fb-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4f7fb-150">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4f7fb-151">命名空间</span><span class="sxs-lookup"><span data-stu-id="4f7fb-151">Namespaces</span></span>

<span data-ttu-id="4f7fb-152">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4f7fb-153">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4f7fb-154">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4f7fb-155">Members</span><span class="sxs-lookup"><span data-stu-id="4f7fb-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4f7fb-156">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="4f7fb-156">ewsUrl: String</span></span>

<span data-ttu-id="4f7fb-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-159">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f7fb-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4f7fb-162">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="4f7fb-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f7fb-165">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-165">Type</span></span>

*   <span data-ttu-id="4f7fb-166">String</span><span class="sxs-lookup"><span data-stu-id="4f7fb-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f7fb-167">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-167">Requirements</span></span>

|<span data-ttu-id="4f7fb-168">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-168">Requirement</span></span>| <span data-ttu-id="4f7fb-169">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-171">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-171">1.0</span></span>|
|[<span data-ttu-id="4f7fb-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-173">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="4f7fb-176">masterCategories： [masterCategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="4f7fb-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="4f7fb-177">获取一个对象，该对象提供用于管理此邮箱上的类别主列表的方法。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-178">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4f7fb-179">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-179">Type</span></span>

*   [<span data-ttu-id="4f7fb-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="4f7fb-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="4f7fb-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-181">Requirements</span></span>

|<span data-ttu-id="4f7fb-182">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-182">Requirement</span></span>| <span data-ttu-id="4f7fb-183">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-185">1.8</span><span class="sxs-lookup"><span data-stu-id="4f7fb-185">1.8</span></span> |
|[<span data-ttu-id="4f7fb-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4f7fb-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="4f7fb-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4f7fb-190">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-190">Example</span></span>

<span data-ttu-id="4f7fb-191">本示例获取此邮箱的类别主列表。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-191">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="4f7fb-192">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="4f7fb-192">restUrl: String</span></span>

<span data-ttu-id="4f7fb-193">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="4f7fb-194">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="4f7fb-195">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="4f7fb-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f7fb-198">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-198">Type</span></span>

*   <span data-ttu-id="4f7fb-199">String</span><span class="sxs-lookup"><span data-stu-id="4f7fb-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f7fb-200">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-200">Requirements</span></span>

|<span data-ttu-id="4f7fb-201">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-201">Requirement</span></span>| <span data-ttu-id="4f7fb-202">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-203">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-204">1.5</span><span class="sxs-lookup"><span data-stu-id="4f7fb-204">1.5</span></span> |
|[<span data-ttu-id="4f7fb-205">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-206">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4f7fb-209">方法</span><span class="sxs-lookup"><span data-stu-id="4f7fb-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4f7fb-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4f7fb-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4f7fb-211">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4f7fb-212">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-213">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-213">Parameters</span></span>

| <span data-ttu-id="4f7fb-214">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-214">Name</span></span> | <span data-ttu-id="4f7fb-215">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-215">Type</span></span> | <span data-ttu-id="4f7fb-216">属性</span><span class="sxs-lookup"><span data-stu-id="4f7fb-216">Attributes</span></span> | <span data-ttu-id="4f7fb-217">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4f7fb-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4f7fb-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4f7fb-219">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4f7fb-220">函数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-220">Function</span></span> || <span data-ttu-id="4f7fb-p105">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4f7fb-224">Object</span><span class="sxs-lookup"><span data-stu-id="4f7fb-224">Object</span></span> | <span data-ttu-id="4f7fb-225">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-225">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-226">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f7fb-227">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-227">Object</span></span> | <span data-ttu-id="4f7fb-228">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-228">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-229">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4f7fb-230">函数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-230">function</span></span>| <span data-ttu-id="4f7fb-231">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-231">&lt;optional&gt;</span></span>|<span data-ttu-id="4f7fb-232">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-233">Requirements</span></span>

|<span data-ttu-id="4f7fb-234">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-234">Requirement</span></span>| <span data-ttu-id="4f7fb-235">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-236">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-237">1.5</span><span class="sxs-lookup"><span data-stu-id="4f7fb-237">1.5</span></span> |
|[<span data-ttu-id="4f7fb-238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-239">ReadItem</span></span> |
|[<span data-ttu-id="4f7fb-240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-242">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-242">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="4f7fb-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4f7fb-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4f7fb-244">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-245">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f7fb-p106">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-248">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-248">Parameters</span></span>

|<span data-ttu-id="4f7fb-249">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-249">Name</span></span>| <span data-ttu-id="4f7fb-250">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-250">Type</span></span>| <span data-ttu-id="4f7fb-251">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f7fb-252">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-252">String</span></span>|<span data-ttu-id="4f7fb-253">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="4f7fb-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4f7fb-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="4f7fb-255">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-256">Requirements</span></span>

|<span data-ttu-id="4f7fb-257">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-257">Requirement</span></span>| <span data-ttu-id="4f7fb-258">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-260">1.3</span><span class="sxs-lookup"><span data-stu-id="4f7fb-260">1.3</span></span>|
|[<span data-ttu-id="4f7fb-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-262">受限</span><span class="sxs-lookup"><span data-stu-id="4f7fb-262">Restricted</span></span>|
|[<span data-ttu-id="4f7fb-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f7fb-265">返回：</span><span class="sxs-lookup"><span data-stu-id="4f7fb-265">Returns:</span></span>

<span data-ttu-id="4f7fb-266">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4f7fb-267">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-267">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="4f7fb-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="4f7fb-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="4f7fb-269">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4f7fb-p107">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4f7fb-p108">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-275">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-275">Parameters</span></span>

|<span data-ttu-id="4f7fb-276">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-276">Name</span></span>| <span data-ttu-id="4f7fb-277">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-277">Type</span></span>| <span data-ttu-id="4f7fb-278">描述</span><span class="sxs-lookup"><span data-stu-id="4f7fb-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4f7fb-279">日期</span><span class="sxs-lookup"><span data-stu-id="4f7fb-279">Date</span></span>|<span data-ttu-id="4f7fb-280">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-281">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-281">Requirements</span></span>

|<span data-ttu-id="4f7fb-282">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-282">Requirement</span></span>| <span data-ttu-id="4f7fb-283">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-285">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-285">1.0</span></span>|
|[<span data-ttu-id="4f7fb-286">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-287">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-288">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-289">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f7fb-290">返回：</span><span class="sxs-lookup"><span data-stu-id="4f7fb-290">Returns:</span></span>

<span data-ttu-id="4f7fb-291">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="4f7fb-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="4f7fb-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4f7fb-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4f7fb-293">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-294">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f7fb-p109">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-297">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-297">Parameters</span></span>

|<span data-ttu-id="4f7fb-298">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-298">Name</span></span>| <span data-ttu-id="4f7fb-299">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-299">Type</span></span>| <span data-ttu-id="4f7fb-300">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f7fb-301">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-301">String</span></span>|<span data-ttu-id="4f7fb-302">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="4f7fb-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4f7fb-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="4f7fb-304">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-305">Requirements</span></span>

|<span data-ttu-id="4f7fb-306">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-306">Requirement</span></span>| <span data-ttu-id="4f7fb-307">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-309">1.3</span><span class="sxs-lookup"><span data-stu-id="4f7fb-309">1.3</span></span>|
|[<span data-ttu-id="4f7fb-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-311">受限</span><span class="sxs-lookup"><span data-stu-id="4f7fb-311">Restricted</span></span>|
|[<span data-ttu-id="4f7fb-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-313">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f7fb-314">返回：</span><span class="sxs-lookup"><span data-stu-id="4f7fb-314">Returns:</span></span>

<span data-ttu-id="4f7fb-315">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4f7fb-316">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-316">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4f7fb-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4f7fb-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4f7fb-318">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4f7fb-319">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-320">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-320">Parameters</span></span>

|<span data-ttu-id="4f7fb-321">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-321">Name</span></span>| <span data-ttu-id="4f7fb-322">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-322">Type</span></span>| <span data-ttu-id="4f7fb-323">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4f7fb-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4f7fb-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="4f7fb-325">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-326">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-326">Requirements</span></span>

|<span data-ttu-id="4f7fb-327">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-327">Requirement</span></span>| <span data-ttu-id="4f7fb-328">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-330">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-330">1.0</span></span>|
|[<span data-ttu-id="4f7fb-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-332">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f7fb-335">返回：</span><span class="sxs-lookup"><span data-stu-id="4f7fb-335">Returns:</span></span>

<span data-ttu-id="4f7fb-336">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-336">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="4f7fb-337">键入：日期</span><span class="sxs-lookup"><span data-stu-id="4f7fb-337">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="4f7fb-338">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-338">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="4f7fb-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4f7fb-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4f7fb-340">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-341">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f7fb-342">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4f7fb-p110">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4f7fb-345">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4f7fb-346">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-347">参数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-347">Parameters</span></span>

|<span data-ttu-id="4f7fb-348">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-348">Name</span></span>| <span data-ttu-id="4f7fb-349">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-349">Type</span></span>| <span data-ttu-id="4f7fb-350">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f7fb-351">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-351">String</span></span>|<span data-ttu-id="4f7fb-352">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-353">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-353">Requirements</span></span>

|<span data-ttu-id="4f7fb-354">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-354">Requirement</span></span>| <span data-ttu-id="4f7fb-355">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-356">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-357">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-357">1.0</span></span>|
|[<span data-ttu-id="4f7fb-358">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-359">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-360">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-361">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-362">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-362">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="4f7fb-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4f7fb-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4f7fb-364">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-365">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f7fb-366">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4f7fb-367">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4f7fb-368">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4f7fb-p111">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-371">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-371">Parameters</span></span>

|<span data-ttu-id="4f7fb-372">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-372">Name</span></span>| <span data-ttu-id="4f7fb-373">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-373">Type</span></span>| <span data-ttu-id="4f7fb-374">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f7fb-375">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-375">String</span></span>|<span data-ttu-id="4f7fb-376">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-377">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-377">Requirements</span></span>

|<span data-ttu-id="4f7fb-378">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-378">Requirement</span></span>| <span data-ttu-id="4f7fb-379">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-380">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-381">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-381">1.0</span></span>|
|[<span data-ttu-id="4f7fb-382">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-383">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-384">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-385">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-386">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-386">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4f7fb-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4f7fb-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4f7fb-388">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-389">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f7fb-p112">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4f7fb-p113">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4f7fb-p114">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4f7fb-397">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-398">参数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-399">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-399">All parameters are optional.</span></span>

|<span data-ttu-id="4f7fb-400">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-400">Name</span></span>| <span data-ttu-id="4f7fb-401">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-401">Type</span></span>| <span data-ttu-id="4f7fb-402">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4f7fb-403">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-403">Object</span></span> | <span data-ttu-id="4f7fb-404">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4f7fb-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4f7fb-p115">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4f7fb-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4f7fb-p116">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4f7fb-411">日期</span><span class="sxs-lookup"><span data-stu-id="4f7fb-411">Date</span></span> | <span data-ttu-id="4f7fb-412">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4f7fb-413">Date</span><span class="sxs-lookup"><span data-stu-id="4f7fb-413">Date</span></span> | <span data-ttu-id="4f7fb-414">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4f7fb-415">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-415">String</span></span> | <span data-ttu-id="4f7fb-p117">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4f7fb-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4f7fb-p118">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4f7fb-421">String</span><span class="sxs-lookup"><span data-stu-id="4f7fb-421">String</span></span> | <span data-ttu-id="4f7fb-p119">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4f7fb-424">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-424">String</span></span> | <span data-ttu-id="4f7fb-p120">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4f7fb-427">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-427">Requirements</span></span>

|<span data-ttu-id="4f7fb-428">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-428">Requirement</span></span>| <span data-ttu-id="4f7fb-429">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-431">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-431">1.0</span></span>|
|[<span data-ttu-id="4f7fb-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-433">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-435">阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-436">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-436">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="4f7fb-437">Office.context.mailbox.displaynewmessageform （参数）</span><span class="sxs-lookup"><span data-stu-id="4f7fb-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="4f7fb-438">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="4f7fb-439">`displayNewMessageForm`方法将打开一个窗体，使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="4f7fb-440">如果指定了参数，则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4f7fb-441">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-442">参数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-443">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-443">All parameters are optional.</span></span>

|<span data-ttu-id="4f7fb-444">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-444">Name</span></span>| <span data-ttu-id="4f7fb-445">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-445">Type</span></span>| <span data-ttu-id="4f7fb-446">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4f7fb-447">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-447">Object</span></span> | <span data-ttu-id="4f7fb-448">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="4f7fb-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4f7fb-450">包含电子邮件地址的字符串数组，或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="4f7fb-451">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="4f7fb-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4f7fb-453">包含电子邮件地址的字符串数组，或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="4f7fb-454">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="4f7fb-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4f7fb-456">包含电子邮件地址的字符串数组，或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="4f7fb-457">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4f7fb-458">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-458">String</span></span> | <span data-ttu-id="4f7fb-459">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-459">A string containing the subject of the message.</span></span> <span data-ttu-id="4f7fb-460">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="4f7fb-461">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-461">String</span></span> | <span data-ttu-id="4f7fb-462">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-462">The HTML body of the message.</span></span> <span data-ttu-id="4f7fb-463">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="4f7fb-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4f7fb-465">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="4f7fb-466">String</span><span class="sxs-lookup"><span data-stu-id="4f7fb-466">String</span></span> | <span data-ttu-id="4f7fb-p127">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="4f7fb-469">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-469">String</span></span> | <span data-ttu-id="4f7fb-470">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="4f7fb-471">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-471">String</span></span> | <span data-ttu-id="4f7fb-p128">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="4f7fb-474">布尔</span><span class="sxs-lookup"><span data-stu-id="4f7fb-474">Boolean</span></span> | <span data-ttu-id="4f7fb-p129">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="4f7fb-477">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-477">String</span></span> | <span data-ttu-id="4f7fb-478">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="4f7fb-479">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="4f7fb-480">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="4f7fb-481">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-481">Requirements</span></span>

|<span data-ttu-id="4f7fb-482">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-482">Requirement</span></span>| <span data-ttu-id="4f7fb-483">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-485">1.6</span><span class="sxs-lookup"><span data-stu-id="4f7fb-485">1.6</span></span> |
|[<span data-ttu-id="4f7fb-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-487">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-489">阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-490">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-490">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="4f7fb-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4f7fb-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="4f7fb-492">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="4f7fb-p131">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-495">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="4f7fb-496">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-496">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="4f7fb-497">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-497">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="4f7fb-498">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-498">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="4f7fb-499">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="4f7fb-499">**REST Tokens**</span></span>

<span data-ttu-id="4f7fb-p133">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="4f7fb-503">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-503">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="4f7fb-504">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="4f7fb-504">**EWS Tokens**</span></span>

<span data-ttu-id="4f7fb-p134">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="4f7fb-507">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-507">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="4f7fb-508">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-508">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4f7fb-509">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以检索附件或项目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-509">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="4f7fb-510">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-510">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-511">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-511">Parameters</span></span>

|<span data-ttu-id="4f7fb-512">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-512">Name</span></span>| <span data-ttu-id="4f7fb-513">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-513">Type</span></span>| <span data-ttu-id="4f7fb-514">属性</span><span class="sxs-lookup"><span data-stu-id="4f7fb-514">Attributes</span></span>| <span data-ttu-id="4f7fb-515">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-515">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="4f7fb-516">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-516">Object</span></span> | <span data-ttu-id="4f7fb-517">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-517">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-518">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-518">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="4f7fb-519">布尔值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-519">Boolean</span></span> |  <span data-ttu-id="4f7fb-520">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-520">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-p136">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f7fb-523">Object</span><span class="sxs-lookup"><span data-stu-id="4f7fb-523">Object</span></span> |  <span data-ttu-id="4f7fb-524">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-524">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-525">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-525">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="4f7fb-526">函数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-526">function</span></span>||<span data-ttu-id="4f7fb-527">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-527">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f7fb-528">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-528">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f7fb-529">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-529">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f7fb-530">错误</span><span class="sxs-lookup"><span data-stu-id="4f7fb-530">Errors</span></span>

|<span data-ttu-id="4f7fb-531">错误代码</span><span class="sxs-lookup"><span data-stu-id="4f7fb-531">Error code</span></span>|<span data-ttu-id="4f7fb-532">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-532">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f7fb-533">请求失败。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-533">The request has failed.</span></span> <span data-ttu-id="4f7fb-534">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-534">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f7fb-535">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-535">The Exchange server returned an error.</span></span> <span data-ttu-id="4f7fb-536">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-536">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f7fb-537">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-537">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f7fb-538">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-538">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-539">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-539">Requirements</span></span>

|<span data-ttu-id="4f7fb-540">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-540">Requirement</span></span>| <span data-ttu-id="4f7fb-541">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-542">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-543">1.5</span><span class="sxs-lookup"><span data-stu-id="4f7fb-543">1.5</span></span> |
|[<span data-ttu-id="4f7fb-544">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-545">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-546">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-547">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-547">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-548">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-548">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4f7fb-549">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f7fb-549">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4f7fb-550">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-550">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4f7fb-p140">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4f7fb-553">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-553">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4f7fb-554">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-554">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="4f7fb-555">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-555">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4f7fb-556">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-556">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="4f7fb-557">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-557">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="4f7fb-558">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-558">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-559">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-559">Parameters</span></span>

|<span data-ttu-id="4f7fb-560">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-560">Name</span></span>| <span data-ttu-id="4f7fb-561">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-561">Type</span></span>| <span data-ttu-id="4f7fb-562">属性</span><span class="sxs-lookup"><span data-stu-id="4f7fb-562">Attributes</span></span>| <span data-ttu-id="4f7fb-563">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-563">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4f7fb-564">function</span><span class="sxs-lookup"><span data-stu-id="4f7fb-564">function</span></span>||<span data-ttu-id="4f7fb-565">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-565">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f7fb-566">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-566">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f7fb-567">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-567">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4f7fb-568">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-568">Object</span></span>| <span data-ttu-id="4f7fb-569">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-569">&lt;optional&gt;</span></span>|<span data-ttu-id="4f7fb-570">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-570">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f7fb-571">错误</span><span class="sxs-lookup"><span data-stu-id="4f7fb-571">Errors</span></span>

|<span data-ttu-id="4f7fb-572">错误代码</span><span class="sxs-lookup"><span data-stu-id="4f7fb-572">Error code</span></span>|<span data-ttu-id="4f7fb-573">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-573">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f7fb-574">请求失败。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-574">The request has failed.</span></span> <span data-ttu-id="4f7fb-575">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-575">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f7fb-576">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-576">The Exchange server returned an error.</span></span> <span data-ttu-id="4f7fb-577">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-577">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f7fb-578">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-578">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f7fb-579">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-579">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-580">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-580">Requirements</span></span>

|<span data-ttu-id="4f7fb-581">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-581">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4f7fb-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-583">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-583">1.0</span></span> | <span data-ttu-id="4f7fb-584">1.3</span><span class="sxs-lookup"><span data-stu-id="4f7fb-584">1.3</span></span> |
|[<span data-ttu-id="4f7fb-585">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-585">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-586">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-586">ReadItem</span></span> | <span data-ttu-id="4f7fb-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-587">ReadItem</span></span> |
|[<span data-ttu-id="4f7fb-588">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-589">阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-589">Read</span></span> | <span data-ttu-id="4f7fb-590">撰写</span><span class="sxs-lookup"><span data-stu-id="4f7fb-590">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="4f7fb-591">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-591">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4f7fb-592">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f7fb-592">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4f7fb-593">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-593">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4f7fb-594">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-594">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-595">参数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-595">Parameters</span></span>

|<span data-ttu-id="4f7fb-596">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-596">Name</span></span>| <span data-ttu-id="4f7fb-597">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-597">Type</span></span>| <span data-ttu-id="4f7fb-598">属性</span><span class="sxs-lookup"><span data-stu-id="4f7fb-598">Attributes</span></span>| <span data-ttu-id="4f7fb-599">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-599">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4f7fb-600">function</span><span class="sxs-lookup"><span data-stu-id="4f7fb-600">function</span></span>||<span data-ttu-id="4f7fb-601">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f7fb-602">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-602">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f7fb-603">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-603">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4f7fb-604">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-604">Object</span></span>| <span data-ttu-id="4f7fb-605">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-605">&lt;optional&gt;</span></span>|<span data-ttu-id="4f7fb-606">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-606">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f7fb-607">错误</span><span class="sxs-lookup"><span data-stu-id="4f7fb-607">Errors</span></span>

|<span data-ttu-id="4f7fb-608">错误代码</span><span class="sxs-lookup"><span data-stu-id="4f7fb-608">Error code</span></span>|<span data-ttu-id="4f7fb-609">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-609">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f7fb-610">请求失败。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-610">The request has failed.</span></span> <span data-ttu-id="4f7fb-611">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-611">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f7fb-612">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-612">The Exchange server returned an error.</span></span> <span data-ttu-id="4f7fb-613">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-613">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f7fb-614">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-614">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f7fb-615">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-615">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-616">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-616">Requirements</span></span>

|<span data-ttu-id="4f7fb-617">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-617">Requirement</span></span>| <span data-ttu-id="4f7fb-618">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-619">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-620">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-620">1.0</span></span>|
|[<span data-ttu-id="4f7fb-621">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-622">ReadItem</span></span>|
|[<span data-ttu-id="4f7fb-623">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-624">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-624">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-625">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-625">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4f7fb-626">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f7fb-626">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4f7fb-627">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-627">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-628">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-628">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4f7fb-629">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="4f7fb-629">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="4f7fb-630">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="4f7fb-630">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4f7fb-631">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-631">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4f7fb-632">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-632">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4f7fb-633">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-633">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4f7fb-634">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-634">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4f7fb-635">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-635">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4f7fb-p150">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4f7fb-638">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-638">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4f7fb-639">版本差异</span><span class="sxs-lookup"><span data-stu-id="4f7fb-639">Version differences</span></span>

<span data-ttu-id="4f7fb-640">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-640">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4f7fb-641">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-641">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="4f7fb-642">您可以使用邮箱. hostName 属性确定您的邮件应用程序是在 web 上的 Outlook 中运行还是在桌面客户端上运行。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-642">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="4f7fb-643">可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-643">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-644">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-644">Parameters</span></span>

|<span data-ttu-id="4f7fb-645">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-645">Name</span></span>| <span data-ttu-id="4f7fb-646">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-646">Type</span></span>| <span data-ttu-id="4f7fb-647">属性</span><span class="sxs-lookup"><span data-stu-id="4f7fb-647">Attributes</span></span>| <span data-ttu-id="4f7fb-648">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-648">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4f7fb-649">字符串</span><span class="sxs-lookup"><span data-stu-id="4f7fb-649">String</span></span>||<span data-ttu-id="4f7fb-650">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-650">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4f7fb-651">函数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-651">function</span></span>||<span data-ttu-id="4f7fb-652">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-652">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f7fb-653">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-653">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4f7fb-654">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-654">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4f7fb-655">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-655">Object</span></span>| <span data-ttu-id="4f7fb-656">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-656">&lt;optional&gt;</span></span>|<span data-ttu-id="4f7fb-657">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-657">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-658">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-658">Requirements</span></span>

|<span data-ttu-id="4f7fb-659">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-659">Requirement</span></span>| <span data-ttu-id="4f7fb-660">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-661">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-662">1.0</span><span class="sxs-lookup"><span data-stu-id="4f7fb-662">1.0</span></span>|
|[<span data-ttu-id="4f7fb-663">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-664">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4f7fb-664">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4f7fb-665">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-666">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-666">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f7fb-667">示例</span><span class="sxs-lookup"><span data-stu-id="4f7fb-667">Example</span></span>

<span data-ttu-id="4f7fb-668">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-668">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4f7fb-669">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4f7fb-669">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4f7fb-670">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-670">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4f7fb-671">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-671">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f7fb-672">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f7fb-672">Parameters</span></span>

| <span data-ttu-id="4f7fb-673">名称</span><span class="sxs-lookup"><span data-stu-id="4f7fb-673">Name</span></span> | <span data-ttu-id="4f7fb-674">类型</span><span class="sxs-lookup"><span data-stu-id="4f7fb-674">Type</span></span> | <span data-ttu-id="4f7fb-675">属性</span><span class="sxs-lookup"><span data-stu-id="4f7fb-675">Attributes</span></span> | <span data-ttu-id="4f7fb-676">说明</span><span class="sxs-lookup"><span data-stu-id="4f7fb-676">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4f7fb-677">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4f7fb-677">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4f7fb-678">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-678">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4f7fb-679">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-679">Object</span></span> | <span data-ttu-id="4f7fb-680">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-680">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-681">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-681">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f7fb-682">对象</span><span class="sxs-lookup"><span data-stu-id="4f7fb-682">Object</span></span> | <span data-ttu-id="4f7fb-683">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-683">&lt;optional&gt;</span></span> | <span data-ttu-id="4f7fb-684">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-684">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4f7fb-685">函数</span><span class="sxs-lookup"><span data-stu-id="4f7fb-685">function</span></span>| <span data-ttu-id="4f7fb-686">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4f7fb-686">&lt;optional&gt;</span></span>|<span data-ttu-id="4f7fb-687">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4f7fb-687">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f7fb-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="4f7fb-688">Requirements</span></span>

|<span data-ttu-id="4f7fb-689">要求</span><span class="sxs-lookup"><span data-stu-id="4f7fb-689">Requirement</span></span>| <span data-ttu-id="4f7fb-690">值</span><span class="sxs-lookup"><span data-stu-id="4f7fb-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f7fb-691">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4f7fb-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f7fb-692">1.5</span><span class="sxs-lookup"><span data-stu-id="4f7fb-692">1.5</span></span> |
|[<span data-ttu-id="4f7fb-693">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4f7fb-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f7fb-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f7fb-694">ReadItem</span></span> |
|[<span data-ttu-id="4f7fb-695">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4f7fb-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f7fb-696">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4f7fb-696">Compose or Read</span></span>|
