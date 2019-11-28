---
title: "\"Context.subname\"-\"邮箱-要求集 1.8\""
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 908eff7b34e63b62fbe250f1a6f810be69b17627
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629214"
---
# <a name="mailbox"></a><span data-ttu-id="7f63c-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="7f63c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="7f63c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="7f63c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="7f63c-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="7f63c-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f63c-105">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-105">Requirements</span></span>

|<span data-ttu-id="7f63c-106">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-106">Requirement</span></span>| <span data-ttu-id="7f63c-107">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-109">1.0</span></span>|
|[<span data-ttu-id="7f63c-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-111">受限</span><span class="sxs-lookup"><span data-stu-id="7f63c-111">Restricted</span></span>|
|[<span data-ttu-id="7f63c-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7f63c-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-114">Members and methods</span></span>

| <span data-ttu-id="7f63c-115">成员</span><span class="sxs-lookup"><span data-stu-id="7f63c-115">Member</span></span> | <span data-ttu-id="7f63c-116">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7f63c-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="7f63c-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="7f63c-118">成员</span><span class="sxs-lookup"><span data-stu-id="7f63c-118">Member</span></span> |
| [<span data-ttu-id="7f63c-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="7f63c-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="7f63c-120">成员</span><span class="sxs-lookup"><span data-stu-id="7f63c-120">Member</span></span> |
| [<span data-ttu-id="7f63c-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="7f63c-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="7f63c-122">成员</span><span class="sxs-lookup"><span data-stu-id="7f63c-122">Member</span></span> |
| [<span data-ttu-id="7f63c-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="7f63c-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="7f63c-124">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-124">Method</span></span> |
| [<span data-ttu-id="7f63c-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="7f63c-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="7f63c-126">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-126">Method</span></span> |
| [<span data-ttu-id="7f63c-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="7f63c-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="7f63c-128">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-128">Method</span></span> |
| [<span data-ttu-id="7f63c-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="7f63c-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="7f63c-130">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-130">Method</span></span> |
| [<span data-ttu-id="7f63c-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="7f63c-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="7f63c-132">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-132">Method</span></span> |
| [<span data-ttu-id="7f63c-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="7f63c-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="7f63c-134">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-134">Method</span></span> |
| [<span data-ttu-id="7f63c-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="7f63c-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="7f63c-136">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-136">Method</span></span> |
| [<span data-ttu-id="7f63c-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="7f63c-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="7f63c-138">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-138">Method</span></span> |
| [<span data-ttu-id="7f63c-139">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="7f63c-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="7f63c-140">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-140">Method</span></span> |
| [<span data-ttu-id="7f63c-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7f63c-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="7f63c-142">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-142">Method</span></span> |
| [<span data-ttu-id="7f63c-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7f63c-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="7f63c-144">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-144">Method</span></span> |
| [<span data-ttu-id="7f63c-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7f63c-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="7f63c-146">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-146">Method</span></span> |
| [<span data-ttu-id="7f63c-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="7f63c-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="7f63c-148">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-148">Method</span></span> |
| [<span data-ttu-id="7f63c-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="7f63c-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="7f63c-150">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7f63c-151">命名空间</span><span class="sxs-lookup"><span data-stu-id="7f63c-151">Namespaces</span></span>

<span data-ttu-id="7f63c-152">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="7f63c-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="7f63c-153">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="7f63c-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="7f63c-154">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="7f63c-155">Members</span><span class="sxs-lookup"><span data-stu-id="7f63c-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="7f63c-156">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="7f63c-156">ewsUrl: String</span></span>

<span data-ttu-id="7f63c-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-159">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="7f63c-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7f63c-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="7f63c-162">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="7f63c-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="7f63c-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="7f63c-165">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-165">Type</span></span>

*   <span data-ttu-id="7f63c-166">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f63c-167">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-167">Requirements</span></span>

|<span data-ttu-id="7f63c-168">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-168">Requirement</span></span>| <span data-ttu-id="7f63c-169">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-171">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-171">1.0</span></span>|
|[<span data-ttu-id="7f63c-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-173">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategoriesviewoutlook-js-18"></a><span data-ttu-id="7f63c-176">masterCategories： [masterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="7f63c-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="7f63c-177">获取一个对象，该对象提供用于管理此邮箱上的类别主列表的方法。</span><span class="sxs-lookup"><span data-stu-id="7f63c-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-178">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="7f63c-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7f63c-179">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-179">Type</span></span>

*   [<span data-ttu-id="7f63c-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="7f63c-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="7f63c-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-181">Requirements</span></span>

|<span data-ttu-id="7f63c-182">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-182">Requirement</span></span>| <span data-ttu-id="7f63c-183">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-185">1.8</span><span class="sxs-lookup"><span data-stu-id="7f63c-185">1.8</span></span> |
|[<span data-ttu-id="7f63c-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="7f63c-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="7f63c-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="7f63c-190">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-190">Example</span></span>

<span data-ttu-id="7f63c-191">本示例获取此邮箱的类别主列表。</span><span class="sxs-lookup"><span data-stu-id="7f63c-191">This example gets the categories master list for this mailbox.</span></span>

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

#### <a name="resturl-string"></a><span data-ttu-id="7f63c-192">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="7f63c-192">restUrl: String</span></span>

<span data-ttu-id="7f63c-193">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="7f63c-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="7f63c-194">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="7f63c-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7f63c-195">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-195">Type</span></span>

*   <span data-ttu-id="7f63c-196">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-196">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f63c-197">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-197">Requirements</span></span>

|<span data-ttu-id="7f63c-198">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-198">Requirement</span></span>| <span data-ttu-id="7f63c-199">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-201">1.5</span><span class="sxs-lookup"><span data-stu-id="7f63c-201">1.5</span></span> |
|[<span data-ttu-id="7f63c-202">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-203">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-204">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-205">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-205">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="7f63c-206">方法</span><span class="sxs-lookup"><span data-stu-id="7f63c-206">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="7f63c-207">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7f63c-207">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="7f63c-208">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="7f63c-208">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="7f63c-209">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="7f63c-209">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-210">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-210">Parameters</span></span>

| <span data-ttu-id="7f63c-211">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-211">Name</span></span> | <span data-ttu-id="7f63c-212">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-212">Type</span></span> | <span data-ttu-id="7f63c-213">属性</span><span class="sxs-lookup"><span data-stu-id="7f63c-213">Attributes</span></span> | <span data-ttu-id="7f63c-214">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-214">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="7f63c-215">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="7f63c-215">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="7f63c-216">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="7f63c-216">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="7f63c-217">函数</span><span class="sxs-lookup"><span data-stu-id="7f63c-217">Function</span></span> || <span data-ttu-id="7f63c-p104">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="7f63c-221">Object</span><span class="sxs-lookup"><span data-stu-id="7f63c-221">Object</span></span> | <span data-ttu-id="7f63c-222">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-222">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-223">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="7f63c-223">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7f63c-224">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-224">Object</span></span> | <span data-ttu-id="7f63c-225">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-225">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-226">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-226">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="7f63c-227">函数</span><span class="sxs-lookup"><span data-stu-id="7f63c-227">function</span></span>| <span data-ttu-id="7f63c-228">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-228">&lt;optional&gt;</span></span>|<span data-ttu-id="7f63c-229">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-229">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-230">Requirements</span></span>

|<span data-ttu-id="7f63c-231">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-231">Requirement</span></span>| <span data-ttu-id="7f63c-232">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-234">1.5</span><span class="sxs-lookup"><span data-stu-id="7f63c-234">1.5</span></span> |
|[<span data-ttu-id="7f63c-235">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-236">ReadItem</span></span> |
|[<span data-ttu-id="7f63c-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-239">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-239">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="7f63c-240">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="7f63c-240">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="7f63c-241">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="7f63c-241">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-242">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="7f63c-242">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7f63c-p105">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-245">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-245">Parameters</span></span>

|<span data-ttu-id="7f63c-246">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-246">Name</span></span>| <span data-ttu-id="7f63c-247">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-247">Type</span></span>| <span data-ttu-id="7f63c-248">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7f63c-249">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-249">String</span></span>|<span data-ttu-id="7f63c-250">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="7f63c-250">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="7f63c-251">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="7f63c-251">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="7f63c-252">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="7f63c-252">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-253">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-253">Requirements</span></span>

|<span data-ttu-id="7f63c-254">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-254">Requirement</span></span>| <span data-ttu-id="7f63c-255">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-256">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-257">1.3</span><span class="sxs-lookup"><span data-stu-id="7f63c-257">1.3</span></span>|
|[<span data-ttu-id="7f63c-258">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-258">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-259">受限</span><span class="sxs-lookup"><span data-stu-id="7f63c-259">Restricted</span></span>|
|[<span data-ttu-id="7f63c-260">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-260">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-261">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-261">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f63c-262">返回：</span><span class="sxs-lookup"><span data-stu-id="7f63c-262">Returns:</span></span>

<span data-ttu-id="7f63c-263">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-263">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7f63c-264">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-264">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-18"></a><span data-ttu-id="7f63c-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="7f63c-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="7f63c-266">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="7f63c-266">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="7f63c-p106">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="7f63c-p107">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-272">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-272">Parameters</span></span>

|<span data-ttu-id="7f63c-273">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-273">Name</span></span>| <span data-ttu-id="7f63c-274">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-274">Type</span></span>| <span data-ttu-id="7f63c-275">描述</span><span class="sxs-lookup"><span data-stu-id="7f63c-275">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="7f63c-276">日期</span><span class="sxs-lookup"><span data-stu-id="7f63c-276">Date</span></span>|<span data-ttu-id="7f63c-277">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-277">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-278">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-278">Requirements</span></span>

|<span data-ttu-id="7f63c-279">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-279">Requirement</span></span>| <span data-ttu-id="7f63c-280">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-281">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-281">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-282">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-282">1.0</span></span>|
|[<span data-ttu-id="7f63c-283">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-283">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-284">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-285">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-285">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-286">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-286">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f63c-287">返回：</span><span class="sxs-lookup"><span data-stu-id="7f63c-287">Returns:</span></span>

<span data-ttu-id="7f63c-288">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="7f63c-288">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="7f63c-289">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="7f63c-289">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="7f63c-290">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="7f63c-290">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-291">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="7f63c-291">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7f63c-p108">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-294">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-294">Parameters</span></span>

|<span data-ttu-id="7f63c-295">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-295">Name</span></span>| <span data-ttu-id="7f63c-296">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-296">Type</span></span>| <span data-ttu-id="7f63c-297">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-297">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7f63c-298">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-298">String</span></span>|<span data-ttu-id="7f63c-299">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="7f63c-299">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="7f63c-300">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="7f63c-300">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="7f63c-301">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="7f63c-301">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-302">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-302">Requirements</span></span>

|<span data-ttu-id="7f63c-303">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-303">Requirement</span></span>| <span data-ttu-id="7f63c-304">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-306">1.3</span><span class="sxs-lookup"><span data-stu-id="7f63c-306">1.3</span></span>|
|[<span data-ttu-id="7f63c-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-308">受限</span><span class="sxs-lookup"><span data-stu-id="7f63c-308">Restricted</span></span>|
|[<span data-ttu-id="7f63c-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-310">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-310">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f63c-311">返回：</span><span class="sxs-lookup"><span data-stu-id="7f63c-311">Returns:</span></span>

<span data-ttu-id="7f63c-312">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-312">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7f63c-313">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-313">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="7f63c-314">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="7f63c-314">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="7f63c-315">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-315">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="7f63c-316">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-316">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-317">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-317">Parameters</span></span>

|<span data-ttu-id="7f63c-318">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-318">Name</span></span>| <span data-ttu-id="7f63c-319">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-319">Type</span></span>| <span data-ttu-id="7f63c-320">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-320">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="7f63c-321">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="7f63c-321">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)|<span data-ttu-id="7f63c-322">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="7f63c-322">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-323">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-323">Requirements</span></span>

|<span data-ttu-id="7f63c-324">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-324">Requirement</span></span>| <span data-ttu-id="7f63c-325">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-326">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-327">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-327">1.0</span></span>|
|[<span data-ttu-id="7f63c-328">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-329">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-330">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-331">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-331">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f63c-332">返回：</span><span class="sxs-lookup"><span data-stu-id="7f63c-332">Returns:</span></span>

<span data-ttu-id="7f63c-333">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-333">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="7f63c-334">键入：日期</span><span class="sxs-lookup"><span data-stu-id="7f63c-334">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="7f63c-335">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-335">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="7f63c-336">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="7f63c-336">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="7f63c-337">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="7f63c-337">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-338">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="7f63c-338">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7f63c-339">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="7f63c-339">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="7f63c-p109">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="7f63c-342">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="7f63c-342">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="7f63c-343">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-343">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-344">参数</span><span class="sxs-lookup"><span data-stu-id="7f63c-344">Parameters</span></span>

|<span data-ttu-id="7f63c-345">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-345">Name</span></span>| <span data-ttu-id="7f63c-346">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-346">Type</span></span>| <span data-ttu-id="7f63c-347">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-347">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7f63c-348">字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-348">String</span></span>|<span data-ttu-id="7f63c-349">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-349">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-350">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-350">Requirements</span></span>

|<span data-ttu-id="7f63c-351">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-351">Requirement</span></span>| <span data-ttu-id="7f63c-352">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-353">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-354">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-354">1.0</span></span>|
|[<span data-ttu-id="7f63c-355">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-356">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-357">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-358">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-358">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-359">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-359">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="7f63c-360">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="7f63c-360">displayMessageForm(itemId)</span></span>

<span data-ttu-id="7f63c-361">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="7f63c-361">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-362">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="7f63c-362">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7f63c-363">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="7f63c-363">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="7f63c-364">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="7f63c-364">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="7f63c-365">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-365">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="7f63c-p110">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-368">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-368">Parameters</span></span>

|<span data-ttu-id="7f63c-369">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-369">Name</span></span>| <span data-ttu-id="7f63c-370">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-370">Type</span></span>| <span data-ttu-id="7f63c-371">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-371">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7f63c-372">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-372">String</span></span>|<span data-ttu-id="7f63c-373">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-373">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-374">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-374">Requirements</span></span>

|<span data-ttu-id="7f63c-375">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-375">Requirement</span></span>| <span data-ttu-id="7f63c-376">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-377">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-378">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-378">1.0</span></span>|
|[<span data-ttu-id="7f63c-379">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-380">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-381">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-382">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-382">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-383">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-383">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="7f63c-384">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="7f63c-384">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="7f63c-385">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="7f63c-385">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-386">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="7f63c-386">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7f63c-p111">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="7f63c-p112">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="7f63c-p113">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="7f63c-394">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="7f63c-394">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-395">参数</span><span class="sxs-lookup"><span data-stu-id="7f63c-395">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-396">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="7f63c-396">All parameters are optional.</span></span>

|<span data-ttu-id="7f63c-397">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-397">Name</span></span>| <span data-ttu-id="7f63c-398">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-398">Type</span></span>| <span data-ttu-id="7f63c-399">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-399">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="7f63c-400">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-400">Object</span></span> | <span data-ttu-id="7f63c-401">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="7f63c-401">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="7f63c-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="7f63c-p114">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="7f63c-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="7f63c-p115">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="7f63c-408">日期</span><span class="sxs-lookup"><span data-stu-id="7f63c-408">Date</span></span> | <span data-ttu-id="7f63c-409">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-409">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="7f63c-410">Date</span><span class="sxs-lookup"><span data-stu-id="7f63c-410">Date</span></span> | <span data-ttu-id="7f63c-411">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-411">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="7f63c-412">字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-412">String</span></span> | <span data-ttu-id="7f63c-p116">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="7f63c-415">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-415">Array.&lt;String&gt;</span></span> | <span data-ttu-id="7f63c-p117">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="7f63c-418">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-418">String</span></span> | <span data-ttu-id="7f63c-p118">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="7f63c-421">字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-421">String</span></span> | <span data-ttu-id="7f63c-p119">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f63c-424">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-424">Requirements</span></span>

|<span data-ttu-id="7f63c-425">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-425">Requirement</span></span>| <span data-ttu-id="7f63c-426">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-427">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-428">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-428">1.0</span></span>|
|[<span data-ttu-id="7f63c-429">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-430">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-431">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-432">阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-432">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-433">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-433">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="7f63c-434">Office.context.mailbox.displaynewmessageform （参数）</span><span class="sxs-lookup"><span data-stu-id="7f63c-434">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="7f63c-435">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="7f63c-435">Displays a form for creating a new message.</span></span>

<span data-ttu-id="7f63c-436">`displayNewMessageForm`方法将打开一个窗体，使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="7f63c-436">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="7f63c-437">如果指定了参数，则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="7f63c-437">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="7f63c-438">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="7f63c-438">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-439">参数</span><span class="sxs-lookup"><span data-stu-id="7f63c-439">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-440">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="7f63c-440">All parameters are optional.</span></span>

|<span data-ttu-id="7f63c-441">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-441">Name</span></span>| <span data-ttu-id="7f63c-442">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-442">Type</span></span>| <span data-ttu-id="7f63c-443">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-443">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="7f63c-444">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-444">Object</span></span> | <span data-ttu-id="7f63c-445">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="7f63c-445">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="7f63c-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="7f63c-447">包含电子邮件地址的字符串数组，或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="7f63c-447">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="7f63c-448">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-448">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="7f63c-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="7f63c-450">包含电子邮件地址的字符串数组，或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="7f63c-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="7f63c-451">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="7f63c-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="7f63c-453">包含电子邮件地址的字符串数组，或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="7f63c-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="7f63c-454">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="7f63c-455">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-455">String</span></span> | <span data-ttu-id="7f63c-456">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="7f63c-456">A string containing the subject of the message.</span></span> <span data-ttu-id="7f63c-457">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-457">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="7f63c-458">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-458">String</span></span> | <span data-ttu-id="7f63c-459">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="7f63c-459">The HTML body of the message.</span></span> <span data-ttu-id="7f63c-460">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="7f63c-460">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="7f63c-461">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-461">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7f63c-462">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="7f63c-462">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="7f63c-463">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-463">String</span></span> | <span data-ttu-id="7f63c-p126">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="7f63c-466">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-466">String</span></span> | <span data-ttu-id="7f63c-467">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-467">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="7f63c-468">String</span><span class="sxs-lookup"><span data-stu-id="7f63c-468">String</span></span> | <span data-ttu-id="7f63c-p127">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="7f63c-471">布尔</span><span class="sxs-lookup"><span data-stu-id="7f63c-471">Boolean</span></span> | <span data-ttu-id="7f63c-p128">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="7f63c-474">字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-474">String</span></span> | <span data-ttu-id="7f63c-475">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="7f63c-475">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="7f63c-476">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="7f63c-476">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="7f63c-477">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="7f63c-477">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="7f63c-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-478">Requirements</span></span>

|<span data-ttu-id="7f63c-479">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-479">Requirement</span></span>| <span data-ttu-id="7f63c-480">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-482">1.6</span><span class="sxs-lookup"><span data-stu-id="7f63c-482">1.6</span></span> |
|[<span data-ttu-id="7f63c-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-484">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-486">阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-487">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-487">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="7f63c-488">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="7f63c-488">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="7f63c-489">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="7f63c-489">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="7f63c-p130">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-492">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="7f63c-492">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="7f63c-493">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="7f63c-493">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="7f63c-494">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-494">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="7f63c-495">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="7f63c-495">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="7f63c-496">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="7f63c-496">**REST Tokens**</span></span>

<span data-ttu-id="7f63c-p132">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="7f63c-500">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="7f63c-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="7f63c-501">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="7f63c-501">**EWS Tokens**</span></span>

<span data-ttu-id="7f63c-p133">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="7f63c-504">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="7f63c-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="7f63c-505">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="7f63c-505">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="7f63c-506">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以检索附件或项目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-506">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="7f63c-507">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="7f63c-507">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-508">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-508">Parameters</span></span>

|<span data-ttu-id="7f63c-509">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-509">Name</span></span>| <span data-ttu-id="7f63c-510">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-510">Type</span></span>| <span data-ttu-id="7f63c-511">属性</span><span class="sxs-lookup"><span data-stu-id="7f63c-511">Attributes</span></span>| <span data-ttu-id="7f63c-512">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-512">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="7f63c-513">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-513">Object</span></span> | <span data-ttu-id="7f63c-514">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-514">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-515">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="7f63c-515">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="7f63c-516">布尔值</span><span class="sxs-lookup"><span data-stu-id="7f63c-516">Boolean</span></span> |  <span data-ttu-id="7f63c-517">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-517">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7f63c-520">Object</span><span class="sxs-lookup"><span data-stu-id="7f63c-520">Object</span></span> |  <span data-ttu-id="7f63c-521">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-521">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-522">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="7f63c-522">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="7f63c-523">函数</span><span class="sxs-lookup"><span data-stu-id="7f63c-523">function</span></span>||<span data-ttu-id="7f63c-524">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-524">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7f63c-525">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="7f63c-525">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="7f63c-526">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-526">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7f63c-527">错误</span><span class="sxs-lookup"><span data-stu-id="7f63c-527">Errors</span></span>

|<span data-ttu-id="7f63c-528">错误代码</span><span class="sxs-lookup"><span data-stu-id="7f63c-528">Error code</span></span>|<span data-ttu-id="7f63c-529">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-529">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="7f63c-530">请求失败。</span><span class="sxs-lookup"><span data-stu-id="7f63c-530">The request has failed.</span></span> <span data-ttu-id="7f63c-531">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="7f63c-531">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="7f63c-532">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="7f63c-532">The Exchange server returned an error.</span></span> <span data-ttu-id="7f63c-533">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-533">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="7f63c-534">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="7f63c-534">The user is no longer connected to the network.</span></span> <span data-ttu-id="7f63c-535">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="7f63c-535">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-536">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-536">Requirements</span></span>

|<span data-ttu-id="7f63c-537">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-537">Requirement</span></span>| <span data-ttu-id="7f63c-538">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-539">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-540">1.5</span><span class="sxs-lookup"><span data-stu-id="7f63c-540">1.5</span></span> |
|[<span data-ttu-id="7f63c-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-542">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-544">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-545">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-545">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="7f63c-546">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7f63c-546">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="7f63c-547">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="7f63c-547">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="7f63c-p139">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="7f63c-550">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="7f63c-550">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="7f63c-551">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-551">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="7f63c-552">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="7f63c-552">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="7f63c-553">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="7f63c-553">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="7f63c-554">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-554">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="7f63c-555">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="7f63c-555">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-556">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-556">Parameters</span></span>

|<span data-ttu-id="7f63c-557">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-557">Name</span></span>| <span data-ttu-id="7f63c-558">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-558">Type</span></span>| <span data-ttu-id="7f63c-559">属性</span><span class="sxs-lookup"><span data-stu-id="7f63c-559">Attributes</span></span>| <span data-ttu-id="7f63c-560">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-560">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7f63c-561">function</span><span class="sxs-lookup"><span data-stu-id="7f63c-561">function</span></span>||<span data-ttu-id="7f63c-562">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-562">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7f63c-563">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="7f63c-563">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="7f63c-564">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-564">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="7f63c-565">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-565">Object</span></span>| <span data-ttu-id="7f63c-566">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-566">&lt;optional&gt;</span></span>|<span data-ttu-id="7f63c-567">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="7f63c-567">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7f63c-568">错误</span><span class="sxs-lookup"><span data-stu-id="7f63c-568">Errors</span></span>

|<span data-ttu-id="7f63c-569">错误代码</span><span class="sxs-lookup"><span data-stu-id="7f63c-569">Error code</span></span>|<span data-ttu-id="7f63c-570">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-570">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="7f63c-571">请求失败。</span><span class="sxs-lookup"><span data-stu-id="7f63c-571">The request has failed.</span></span> <span data-ttu-id="7f63c-572">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="7f63c-572">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="7f63c-573">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="7f63c-573">The Exchange server returned an error.</span></span> <span data-ttu-id="7f63c-574">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-574">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="7f63c-575">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="7f63c-575">The user is no longer connected to the network.</span></span> <span data-ttu-id="7f63c-576">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="7f63c-576">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-577">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-577">Requirements</span></span>

|<span data-ttu-id="7f63c-578">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-578">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="7f63c-579">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-579">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-580">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-580">1.0</span></span> | <span data-ttu-id="7f63c-581">1.3</span><span class="sxs-lookup"><span data-stu-id="7f63c-581">1.3</span></span> |
|[<span data-ttu-id="7f63c-582">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-583">ReadItem</span></span> | <span data-ttu-id="7f63c-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-584">ReadItem</span></span> |
|[<span data-ttu-id="7f63c-585">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-586">阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-586">Read</span></span> | <span data-ttu-id="7f63c-587">撰写</span><span class="sxs-lookup"><span data-stu-id="7f63c-587">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="7f63c-588">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-588">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="7f63c-589">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7f63c-589">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="7f63c-590">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="7f63c-590">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="7f63c-591">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="7f63c-591">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-592">参数</span><span class="sxs-lookup"><span data-stu-id="7f63c-592">Parameters</span></span>

|<span data-ttu-id="7f63c-593">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-593">Name</span></span>| <span data-ttu-id="7f63c-594">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-594">Type</span></span>| <span data-ttu-id="7f63c-595">属性</span><span class="sxs-lookup"><span data-stu-id="7f63c-595">Attributes</span></span>| <span data-ttu-id="7f63c-596">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-596">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7f63c-597">function</span><span class="sxs-lookup"><span data-stu-id="7f63c-597">function</span></span>||<span data-ttu-id="7f63c-598">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7f63c-599">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="7f63c-599">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="7f63c-600">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-600">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="7f63c-601">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-601">Object</span></span>| <span data-ttu-id="7f63c-602">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-602">&lt;optional&gt;</span></span>|<span data-ttu-id="7f63c-603">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="7f63c-603">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7f63c-604">错误</span><span class="sxs-lookup"><span data-stu-id="7f63c-604">Errors</span></span>

|<span data-ttu-id="7f63c-605">错误代码</span><span class="sxs-lookup"><span data-stu-id="7f63c-605">Error code</span></span>|<span data-ttu-id="7f63c-606">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-606">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="7f63c-607">请求失败。</span><span class="sxs-lookup"><span data-stu-id="7f63c-607">The request has failed.</span></span> <span data-ttu-id="7f63c-608">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="7f63c-608">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="7f63c-609">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="7f63c-609">The Exchange server returned an error.</span></span> <span data-ttu-id="7f63c-610">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-610">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="7f63c-611">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="7f63c-611">The user is no longer connected to the network.</span></span> <span data-ttu-id="7f63c-612">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="7f63c-612">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-613">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-613">Requirements</span></span>

|<span data-ttu-id="7f63c-614">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-614">Requirement</span></span>| <span data-ttu-id="7f63c-615">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-616">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-617">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-617">1.0</span></span>|
|[<span data-ttu-id="7f63c-618">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-619">ReadItem</span></span>|
|[<span data-ttu-id="7f63c-620">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-621">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-622">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-622">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="7f63c-623">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7f63c-623">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="7f63c-624">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="7f63c-624">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-625">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="7f63c-625">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="7f63c-626">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="7f63c-626">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="7f63c-627">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="7f63c-627">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="7f63c-628">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="7f63c-628">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="7f63c-629">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="7f63c-629">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="7f63c-630">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="7f63c-630">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="7f63c-631">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="7f63c-631">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="7f63c-632">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="7f63c-632">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="7f63c-p149">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="7f63c-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="7f63c-635">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="7f63c-635">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="7f63c-636">版本差异</span><span class="sxs-lookup"><span data-stu-id="7f63c-636">Version differences</span></span>

<span data-ttu-id="7f63c-637">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="7f63c-637">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="7f63c-638">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。</span><span class="sxs-lookup"><span data-stu-id="7f63c-638">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="7f63c-639">您可以使用邮箱. hostName 属性确定您的邮件应用程序是在 web 上的 Outlook 中运行还是在桌面客户端上运行。</span><span class="sxs-lookup"><span data-stu-id="7f63c-639">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="7f63c-640">可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="7f63c-640">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-641">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-641">Parameters</span></span>

|<span data-ttu-id="7f63c-642">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-642">Name</span></span>| <span data-ttu-id="7f63c-643">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-643">Type</span></span>| <span data-ttu-id="7f63c-644">属性</span><span class="sxs-lookup"><span data-stu-id="7f63c-644">Attributes</span></span>| <span data-ttu-id="7f63c-645">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-645">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7f63c-646">字符串</span><span class="sxs-lookup"><span data-stu-id="7f63c-646">String</span></span>||<span data-ttu-id="7f63c-647">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="7f63c-647">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="7f63c-648">函数</span><span class="sxs-lookup"><span data-stu-id="7f63c-648">function</span></span>||<span data-ttu-id="7f63c-649">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7f63c-650">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="7f63c-650">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="7f63c-651">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="7f63c-651">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="7f63c-652">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-652">Object</span></span>| <span data-ttu-id="7f63c-653">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-653">&lt;optional&gt;</span></span>|<span data-ttu-id="7f63c-654">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="7f63c-654">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-655">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-655">Requirements</span></span>

|<span data-ttu-id="7f63c-656">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-656">Requirement</span></span>| <span data-ttu-id="7f63c-657">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-657">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-658">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-658">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-659">1.0</span><span class="sxs-lookup"><span data-stu-id="7f63c-659">1.0</span></span>|
|[<span data-ttu-id="7f63c-660">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-660">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-661">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="7f63c-661">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="7f63c-662">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-662">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-663">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-663">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f63c-664">示例</span><span class="sxs-lookup"><span data-stu-id="7f63c-664">Example</span></span>

<span data-ttu-id="7f63c-665">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="7f63c-665">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="7f63c-666">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7f63c-666">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="7f63c-667">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="7f63c-667">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="7f63c-668">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="7f63c-668">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f63c-669">Parameters</span><span class="sxs-lookup"><span data-stu-id="7f63c-669">Parameters</span></span>

| <span data-ttu-id="7f63c-670">名称</span><span class="sxs-lookup"><span data-stu-id="7f63c-670">Name</span></span> | <span data-ttu-id="7f63c-671">类型</span><span class="sxs-lookup"><span data-stu-id="7f63c-671">Type</span></span> | <span data-ttu-id="7f63c-672">属性</span><span class="sxs-lookup"><span data-stu-id="7f63c-672">Attributes</span></span> | <span data-ttu-id="7f63c-673">说明</span><span class="sxs-lookup"><span data-stu-id="7f63c-673">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="7f63c-674">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="7f63c-674">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="7f63c-675">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="7f63c-675">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="7f63c-676">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-676">Object</span></span> | <span data-ttu-id="7f63c-677">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-677">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-678">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="7f63c-678">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7f63c-679">对象</span><span class="sxs-lookup"><span data-stu-id="7f63c-679">Object</span></span> | <span data-ttu-id="7f63c-680">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-680">&lt;optional&gt;</span></span> | <span data-ttu-id="7f63c-681">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="7f63c-681">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="7f63c-682">函数</span><span class="sxs-lookup"><span data-stu-id="7f63c-682">function</span></span>| <span data-ttu-id="7f63c-683">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="7f63c-683">&lt;optional&gt;</span></span>|<span data-ttu-id="7f63c-684">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="7f63c-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f63c-685">Requirements</span><span class="sxs-lookup"><span data-stu-id="7f63c-685">Requirements</span></span>

|<span data-ttu-id="7f63c-686">要求</span><span class="sxs-lookup"><span data-stu-id="7f63c-686">Requirement</span></span>| <span data-ttu-id="7f63c-687">值</span><span class="sxs-lookup"><span data-stu-id="7f63c-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f63c-688">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f63c-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f63c-689">1.5</span><span class="sxs-lookup"><span data-stu-id="7f63c-689">1.5</span></span> |
|[<span data-ttu-id="7f63c-690">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7f63c-690">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f63c-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f63c-691">ReadItem</span></span> |
|[<span data-ttu-id="7f63c-692">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f63c-692">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f63c-693">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f63c-693">Compose or Read</span></span>|
