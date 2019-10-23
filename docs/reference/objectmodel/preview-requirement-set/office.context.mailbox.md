---
title: "\"Context\"-\"邮箱-预览要求集\""
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 29922c9e05cc0380f1e54a16f3350c578d9e4cee
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627067"
---
# <a name="mailbox"></a><span data-ttu-id="3a494-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="3a494-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="3a494-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="3a494-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="3a494-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3a494-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a494-105">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-105">Requirements</span></span>

|<span data-ttu-id="3a494-106">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-106">Requirement</span></span>| <span data-ttu-id="3a494-107">值</span><span class="sxs-lookup"><span data-stu-id="3a494-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-109">1.0</span></span>|
|[<span data-ttu-id="3a494-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-111">受限</span><span class="sxs-lookup"><span data-stu-id="3a494-111">Restricted</span></span>|
|[<span data-ttu-id="3a494-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3a494-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="3a494-114">Members and methods</span></span>

| <span data-ttu-id="3a494-115">成员</span><span class="sxs-lookup"><span data-stu-id="3a494-115">Member</span></span> | <span data-ttu-id="3a494-116">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3a494-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="3a494-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="3a494-118">成员</span><span class="sxs-lookup"><span data-stu-id="3a494-118">Member</span></span> |
| [<span data-ttu-id="3a494-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="3a494-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="3a494-120">成员</span><span class="sxs-lookup"><span data-stu-id="3a494-120">Member</span></span> |
| [<span data-ttu-id="3a494-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="3a494-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="3a494-122">成员</span><span class="sxs-lookup"><span data-stu-id="3a494-122">Member</span></span> |
| [<span data-ttu-id="3a494-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3a494-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3a494-124">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-124">Method</span></span> |
| [<span data-ttu-id="3a494-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="3a494-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="3a494-126">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-126">Method</span></span> |
| [<span data-ttu-id="3a494-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3a494-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="3a494-128">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-128">Method</span></span> |
| [<span data-ttu-id="3a494-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="3a494-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="3a494-130">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-130">Method</span></span> |
| [<span data-ttu-id="3a494-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="3a494-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="3a494-132">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-132">Method</span></span> |
| [<span data-ttu-id="3a494-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3a494-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="3a494-134">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-134">Method</span></span> |
| [<span data-ttu-id="3a494-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="3a494-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="3a494-136">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-136">Method</span></span> |
| [<span data-ttu-id="3a494-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3a494-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="3a494-138">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-138">Method</span></span> |
| [<span data-ttu-id="3a494-139">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="3a494-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="3a494-140">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-140">Method</span></span> |
| [<span data-ttu-id="3a494-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3a494-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="3a494-142">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-142">Method</span></span> |
| [<span data-ttu-id="3a494-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3a494-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="3a494-144">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-144">Method</span></span> |
| [<span data-ttu-id="3a494-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3a494-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="3a494-146">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-146">Method</span></span> |
| [<span data-ttu-id="3a494-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="3a494-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="3a494-148">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-148">Method</span></span> |
| [<span data-ttu-id="3a494-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3a494-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="3a494-150">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3a494-151">命名空间</span><span class="sxs-lookup"><span data-stu-id="3a494-151">Namespaces</span></span>

<span data-ttu-id="3a494-152">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="3a494-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="3a494-153">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="3a494-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="3a494-154">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="3a494-155">Members</span><span class="sxs-lookup"><span data-stu-id="3a494-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="3a494-156">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="3a494-156">ewsUrl: String</span></span>

<span data-ttu-id="3a494-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3a494-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-159">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3a494-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a494-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3a494-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3a494-162">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="3a494-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="3a494-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3a494-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3a494-165">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-165">Type</span></span>

*   <span data-ttu-id="3a494-166">String</span><span class="sxs-lookup"><span data-stu-id="3a494-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a494-167">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-167">Requirements</span></span>

|<span data-ttu-id="3a494-168">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-168">Requirement</span></span>| <span data-ttu-id="3a494-169">值</span><span class="sxs-lookup"><span data-stu-id="3a494-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-171">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-171">1.0</span></span>|
|[<span data-ttu-id="3a494-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-173">ReadItem</span></span>|
|[<span data-ttu-id="3a494-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="3a494-176">masterCategories： [masterCategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="3a494-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="3a494-177">获取一个对象，该对象提供用于管理此邮箱上的类别主列表的方法。</span><span class="sxs-lookup"><span data-stu-id="3a494-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-178">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3a494-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="3a494-179">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-179">Type</span></span>

*   [<span data-ttu-id="3a494-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="3a494-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="3a494-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-181">Requirements</span></span>

|<span data-ttu-id="3a494-182">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-182">Requirement</span></span>| <span data-ttu-id="3a494-183">值</span><span class="sxs-lookup"><span data-stu-id="3a494-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-185">预览</span><span class="sxs-lookup"><span data-stu-id="3a494-185">Preview</span></span> |
|[<span data-ttu-id="3a494-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3a494-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="3a494-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="3a494-190">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-190">Example</span></span>

<span data-ttu-id="3a494-191">本示例获取此邮箱的类别主列表。</span><span class="sxs-lookup"><span data-stu-id="3a494-191">This example gets the categories master list for this mailbox.</span></span>

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

#### <a name="resturl-string"></a><span data-ttu-id="3a494-192">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="3a494-192">restUrl: String</span></span>

<span data-ttu-id="3a494-193">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="3a494-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="3a494-194">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="3a494-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="3a494-195">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="3a494-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="3a494-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3a494-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3a494-198">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-198">Type</span></span>

*   <span data-ttu-id="3a494-199">String</span><span class="sxs-lookup"><span data-stu-id="3a494-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a494-200">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-200">Requirements</span></span>

|<span data-ttu-id="3a494-201">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-201">Requirement</span></span>| <span data-ttu-id="3a494-202">值</span><span class="sxs-lookup"><span data-stu-id="3a494-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-203">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-204">1.5</span><span class="sxs-lookup"><span data-stu-id="3a494-204">1.5</span></span> |
|[<span data-ttu-id="3a494-205">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-206">ReadItem</span></span>|
|[<span data-ttu-id="3a494-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3a494-209">方法</span><span class="sxs-lookup"><span data-stu-id="3a494-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3a494-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3a494-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3a494-211">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3a494-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3a494-212">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="3a494-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-213">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-213">Parameters</span></span>

| <span data-ttu-id="3a494-214">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-214">Name</span></span> | <span data-ttu-id="3a494-215">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-215">Type</span></span> | <span data-ttu-id="3a494-216">属性</span><span class="sxs-lookup"><span data-stu-id="3a494-216">Attributes</span></span> | <span data-ttu-id="3a494-217">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3a494-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3a494-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3a494-219">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3a494-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3a494-220">函数</span><span class="sxs-lookup"><span data-stu-id="3a494-220">Function</span></span> || <span data-ttu-id="3a494-p105">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="3a494-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3a494-224">Object</span><span class="sxs-lookup"><span data-stu-id="3a494-224">Object</span></span> | <span data-ttu-id="3a494-225">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-225">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-226">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3a494-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3a494-227">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-227">Object</span></span> | <span data-ttu-id="3a494-228">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-228">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-229">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3a494-230">函数</span><span class="sxs-lookup"><span data-stu-id="3a494-230">function</span></span>| <span data-ttu-id="3a494-231">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-231">&lt;optional&gt;</span></span>|<span data-ttu-id="3a494-232">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a494-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-233">Requirements</span></span>

|<span data-ttu-id="3a494-234">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-234">Requirement</span></span>| <span data-ttu-id="3a494-235">值</span><span class="sxs-lookup"><span data-stu-id="3a494-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-236">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-237">1.5</span><span class="sxs-lookup"><span data-stu-id="3a494-237">1.5</span></span> |
|[<span data-ttu-id="3a494-238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-239">ReadItem</span></span> |
|[<span data-ttu-id="3a494-240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-242">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-242">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="3a494-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3a494-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3a494-244">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="3a494-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-245">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a494-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a494-p106">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="3a494-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-248">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-248">Parameters</span></span>

|<span data-ttu-id="3a494-249">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-249">Name</span></span>| <span data-ttu-id="3a494-250">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-250">Type</span></span>| <span data-ttu-id="3a494-251">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a494-252">String</span><span class="sxs-lookup"><span data-stu-id="3a494-252">String</span></span>|<span data-ttu-id="3a494-253">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="3a494-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="3a494-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3a494-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="3a494-255">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="3a494-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-256">Requirements</span></span>

|<span data-ttu-id="3a494-257">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-257">Requirement</span></span>| <span data-ttu-id="3a494-258">值</span><span class="sxs-lookup"><span data-stu-id="3a494-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-260">1.3</span><span class="sxs-lookup"><span data-stu-id="3a494-260">1.3</span></span>|
|[<span data-ttu-id="3a494-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-262">受限</span><span class="sxs-lookup"><span data-stu-id="3a494-262">Restricted</span></span>|
|[<span data-ttu-id="3a494-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a494-265">返回：</span><span class="sxs-lookup"><span data-stu-id="3a494-265">Returns:</span></span>

<span data-ttu-id="3a494-266">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3a494-267">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-267">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="3a494-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="3a494-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="3a494-269">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="3a494-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="3a494-p107">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="3a494-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="3a494-p108">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-275">Parameters</span><span class="sxs-lookup"><span data-stu-id="3a494-275">Parameters</span></span>

|<span data-ttu-id="3a494-276">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-276">Name</span></span>| <span data-ttu-id="3a494-277">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-277">Type</span></span>| <span data-ttu-id="3a494-278">描述</span><span class="sxs-lookup"><span data-stu-id="3a494-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="3a494-279">日期</span><span class="sxs-lookup"><span data-stu-id="3a494-279">Date</span></span>|<span data-ttu-id="3a494-280">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="3a494-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-281">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-281">Requirements</span></span>

|<span data-ttu-id="3a494-282">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-282">Requirement</span></span>| <span data-ttu-id="3a494-283">值</span><span class="sxs-lookup"><span data-stu-id="3a494-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-285">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-285">1.0</span></span>|
|[<span data-ttu-id="3a494-286">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-287">ReadItem</span></span>|
|[<span data-ttu-id="3a494-288">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-289">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a494-290">返回：</span><span class="sxs-lookup"><span data-stu-id="3a494-290">Returns:</span></span>

<span data-ttu-id="3a494-291">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="3a494-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="3a494-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3a494-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3a494-293">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="3a494-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-294">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a494-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a494-p109">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="3a494-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-297">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-297">Parameters</span></span>

|<span data-ttu-id="3a494-298">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-298">Name</span></span>| <span data-ttu-id="3a494-299">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-299">Type</span></span>| <span data-ttu-id="3a494-300">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a494-301">String</span><span class="sxs-lookup"><span data-stu-id="3a494-301">String</span></span>|<span data-ttu-id="3a494-302">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="3a494-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="3a494-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3a494-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="3a494-304">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="3a494-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-305">Requirements</span></span>

|<span data-ttu-id="3a494-306">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-306">Requirement</span></span>| <span data-ttu-id="3a494-307">值</span><span class="sxs-lookup"><span data-stu-id="3a494-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-309">1.3</span><span class="sxs-lookup"><span data-stu-id="3a494-309">1.3</span></span>|
|[<span data-ttu-id="3a494-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-311">受限</span><span class="sxs-lookup"><span data-stu-id="3a494-311">Restricted</span></span>|
|[<span data-ttu-id="3a494-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-313">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a494-314">返回：</span><span class="sxs-lookup"><span data-stu-id="3a494-314">Returns:</span></span>

<span data-ttu-id="3a494-315">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3a494-316">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-316">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="3a494-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="3a494-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="3a494-318">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="3a494-319">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-320">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-320">Parameters</span></span>

|<span data-ttu-id="3a494-321">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-321">Name</span></span>| <span data-ttu-id="3a494-322">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-322">Type</span></span>| <span data-ttu-id="3a494-323">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="3a494-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3a494-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="3a494-325">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="3a494-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-326">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-326">Requirements</span></span>

|<span data-ttu-id="3a494-327">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-327">Requirement</span></span>| <span data-ttu-id="3a494-328">值</span><span class="sxs-lookup"><span data-stu-id="3a494-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-330">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-330">1.0</span></span>|
|[<span data-ttu-id="3a494-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-332">ReadItem</span></span>|
|[<span data-ttu-id="3a494-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a494-335">返回：</span><span class="sxs-lookup"><span data-stu-id="3a494-335">Returns:</span></span>

<span data-ttu-id="3a494-336">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-336">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="3a494-337">键入：日期</span><span class="sxs-lookup"><span data-stu-id="3a494-337">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="3a494-338">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-338">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="3a494-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3a494-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="3a494-340">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="3a494-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-341">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a494-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a494-342">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="3a494-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3a494-p110">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="3a494-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="3a494-345">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a494-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="3a494-346">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="3a494-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-347">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-347">Parameters</span></span>

|<span data-ttu-id="3a494-348">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-348">Name</span></span>| <span data-ttu-id="3a494-349">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-349">Type</span></span>| <span data-ttu-id="3a494-350">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a494-351">字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-351">String</span></span>|<span data-ttu-id="3a494-352">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="3a494-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-353">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-353">Requirements</span></span>

|<span data-ttu-id="3a494-354">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-354">Requirement</span></span>| <span data-ttu-id="3a494-355">值</span><span class="sxs-lookup"><span data-stu-id="3a494-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-356">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-357">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-357">1.0</span></span>|
|[<span data-ttu-id="3a494-358">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-359">ReadItem</span></span>|
|[<span data-ttu-id="3a494-360">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-361">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-362">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-362">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="3a494-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3a494-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="3a494-364">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="3a494-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-365">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a494-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a494-366">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="3a494-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3a494-367">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a494-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="3a494-368">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="3a494-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="3a494-p111">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="3a494-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-371">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-371">Parameters</span></span>

|<span data-ttu-id="3a494-372">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-372">Name</span></span>| <span data-ttu-id="3a494-373">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-373">Type</span></span>| <span data-ttu-id="3a494-374">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a494-375">String</span><span class="sxs-lookup"><span data-stu-id="3a494-375">String</span></span>|<span data-ttu-id="3a494-376">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="3a494-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-377">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-377">Requirements</span></span>

|<span data-ttu-id="3a494-378">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-378">Requirement</span></span>| <span data-ttu-id="3a494-379">值</span><span class="sxs-lookup"><span data-stu-id="3a494-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-380">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-381">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-381">1.0</span></span>|
|[<span data-ttu-id="3a494-382">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-383">ReadItem</span></span>|
|[<span data-ttu-id="3a494-384">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-385">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-386">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-386">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="3a494-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3a494-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="3a494-388">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="3a494-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-389">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a494-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a494-p112">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="3a494-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3a494-p113">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="3a494-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="3a494-p114">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="3a494-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="3a494-397">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="3a494-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-398">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-399">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="3a494-399">All parameters are optional.</span></span>

|<span data-ttu-id="3a494-400">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-400">Name</span></span>| <span data-ttu-id="3a494-401">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-401">Type</span></span>| <span data-ttu-id="3a494-402">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3a494-403">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-403">Object</span></span> | <span data-ttu-id="3a494-404">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="3a494-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="3a494-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a494-p115">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a494-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="3a494-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a494-p116">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a494-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="3a494-411">日期</span><span class="sxs-lookup"><span data-stu-id="3a494-411">Date</span></span> | <span data-ttu-id="3a494-412">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="3a494-413">Date</span><span class="sxs-lookup"><span data-stu-id="3a494-413">Date</span></span> | <span data-ttu-id="3a494-414">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="3a494-415">字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-415">String</span></span> | <span data-ttu-id="3a494-p117">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a494-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="3a494-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="3a494-p118">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a494-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3a494-421">String</span><span class="sxs-lookup"><span data-stu-id="3a494-421">String</span></span> | <span data-ttu-id="3a494-p119">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a494-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="3a494-424">字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-424">String</span></span> | <span data-ttu-id="3a494-p120">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3a494-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a494-427">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-427">Requirements</span></span>

|<span data-ttu-id="3a494-428">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-428">Requirement</span></span>| <span data-ttu-id="3a494-429">值</span><span class="sxs-lookup"><span data-stu-id="3a494-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-431">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-431">1.0</span></span>|
|[<span data-ttu-id="3a494-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-433">ReadItem</span></span>|
|[<span data-ttu-id="3a494-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-435">阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-436">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="3a494-437">Office.context.mailbox.displaynewmessageform （参数）</span><span class="sxs-lookup"><span data-stu-id="3a494-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="3a494-438">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a494-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="3a494-439">`displayNewMessageForm`方法将打开一个窗体，使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="3a494-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="3a494-440">如果指定了参数，则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="3a494-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3a494-441">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="3a494-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-442">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-443">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="3a494-443">All parameters are optional.</span></span>

|<span data-ttu-id="3a494-444">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-444">Name</span></span>| <span data-ttu-id="3a494-445">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-445">Type</span></span>| <span data-ttu-id="3a494-446">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3a494-447">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-447">Object</span></span> | <span data-ttu-id="3a494-448">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="3a494-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="3a494-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a494-450">包含电子邮件地址的字符串数组，或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3a494-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="3a494-451">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a494-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="3a494-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a494-453">包含电子邮件地址的字符串数组，或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3a494-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="3a494-454">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a494-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="3a494-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a494-456">包含电子邮件地址的字符串数组，或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3a494-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="3a494-457">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a494-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3a494-458">String</span><span class="sxs-lookup"><span data-stu-id="3a494-458">String</span></span> | <span data-ttu-id="3a494-459">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="3a494-459">A string containing the subject of the message.</span></span> <span data-ttu-id="3a494-460">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a494-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="3a494-461">String</span><span class="sxs-lookup"><span data-stu-id="3a494-461">String</span></span> | <span data-ttu-id="3a494-462">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="3a494-462">The HTML body of the message.</span></span> <span data-ttu-id="3a494-463">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3a494-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="3a494-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3a494-465">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3a494-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="3a494-466">String</span><span class="sxs-lookup"><span data-stu-id="3a494-466">String</span></span> | <span data-ttu-id="3a494-p127">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3a494-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="3a494-469">String</span><span class="sxs-lookup"><span data-stu-id="3a494-469">String</span></span> | <span data-ttu-id="3a494-470">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a494-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="3a494-471">String</span><span class="sxs-lookup"><span data-stu-id="3a494-471">String</span></span> | <span data-ttu-id="3a494-p128">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3a494-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="3a494-474">布尔</span><span class="sxs-lookup"><span data-stu-id="3a494-474">Boolean</span></span> | <span data-ttu-id="3a494-p129">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="3a494-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="3a494-477">字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-477">String</span></span> | <span data-ttu-id="3a494-478">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="3a494-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="3a494-479">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="3a494-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="3a494-480">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a494-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="3a494-481">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-481">Requirements</span></span>

|<span data-ttu-id="3a494-482">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-482">Requirement</span></span>| <span data-ttu-id="3a494-483">值</span><span class="sxs-lookup"><span data-stu-id="3a494-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-485">1.6</span><span class="sxs-lookup"><span data-stu-id="3a494-485">1.6</span></span> |
|[<span data-ttu-id="3a494-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-487">ReadItem</span></span>|
|[<span data-ttu-id="3a494-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-489">阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-490">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-490">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="3a494-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3a494-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="3a494-492">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="3a494-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="3a494-p131">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="3a494-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-495">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="3a494-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="3a494-496">在读取`getCallbackTokenAsync`模式下调用方法需要**ReadItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="3a494-496">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="3a494-497">在`getCallbackTokenAsync`撰写模式下调用需要您保存项目。</span><span class="sxs-lookup"><span data-stu-id="3a494-497">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="3a494-498">该[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)方法需要**ReadWriteItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="3a494-498">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="3a494-499">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="3a494-499">**REST Tokens**</span></span>

<span data-ttu-id="3a494-p133">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="3a494-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="3a494-503">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="3a494-503">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="3a494-504">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="3a494-504">**EWS Tokens**</span></span>

<span data-ttu-id="3a494-p134">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="3a494-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="3a494-507">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="3a494-507">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="3a494-508">您可以将令牌和附件标识符或项目标识符同时传递给第三方系统。</span><span class="sxs-lookup"><span data-stu-id="3a494-508">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="3a494-509">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以检索附件或项目。</span><span class="sxs-lookup"><span data-stu-id="3a494-509">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="3a494-510">例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3a494-510">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-511">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-511">Parameters</span></span>

|<span data-ttu-id="3a494-512">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-512">Name</span></span>| <span data-ttu-id="3a494-513">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-513">Type</span></span>| <span data-ttu-id="3a494-514">属性</span><span class="sxs-lookup"><span data-stu-id="3a494-514">Attributes</span></span>| <span data-ttu-id="3a494-515">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-515">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="3a494-516">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-516">Object</span></span> | <span data-ttu-id="3a494-517">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-517">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-518">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3a494-518">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="3a494-519">布尔值</span><span class="sxs-lookup"><span data-stu-id="3a494-519">Boolean</span></span> |  <span data-ttu-id="3a494-520">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-520">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-p136">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="3a494-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3a494-523">Object</span><span class="sxs-lookup"><span data-stu-id="3a494-523">Object</span></span> |  <span data-ttu-id="3a494-524">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-524">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-525">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a494-525">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="3a494-526">函数</span><span class="sxs-lookup"><span data-stu-id="3a494-526">function</span></span>||<span data-ttu-id="3a494-527">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a494-527">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a494-528">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a494-528">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3a494-529">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-529">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3a494-530">错误</span><span class="sxs-lookup"><span data-stu-id="3a494-530">Errors</span></span>

|<span data-ttu-id="3a494-531">错误代码</span><span class="sxs-lookup"><span data-stu-id="3a494-531">Error code</span></span>|<span data-ttu-id="3a494-532">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-532">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3a494-533">请求失败。</span><span class="sxs-lookup"><span data-stu-id="3a494-533">The request has failed.</span></span> <span data-ttu-id="3a494-534">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="3a494-534">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3a494-535">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="3a494-535">The Exchange server returned an error.</span></span> <span data-ttu-id="3a494-536">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-536">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3a494-537">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="3a494-537">The user is no longer connected to the network.</span></span> <span data-ttu-id="3a494-538">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="3a494-538">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-539">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-539">Requirements</span></span>

|<span data-ttu-id="3a494-540">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-540">Requirement</span></span>| <span data-ttu-id="3a494-541">值</span><span class="sxs-lookup"><span data-stu-id="3a494-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-542">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-543">1.5</span><span class="sxs-lookup"><span data-stu-id="3a494-543">1.5</span></span> |
|[<span data-ttu-id="3a494-544">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-545">ReadItem</span></span>|
|[<span data-ttu-id="3a494-546">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-547">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-547">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-548">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-548">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="3a494-549">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a494-549">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3a494-550">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="3a494-550">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="3a494-p140">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="3a494-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="3a494-553">您可以将令牌和附件标识符或项目标识符同时传递给第三方系统。</span><span class="sxs-lookup"><span data-stu-id="3a494-553">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="3a494-554">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="3a494-554">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="3a494-555">例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3a494-555">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3a494-556">在读取`getCallbackTokenAsync`模式下调用方法需要**ReadItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="3a494-556">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="3a494-557">在`getCallbackTokenAsync`撰写模式下调用需要您保存项目。</span><span class="sxs-lookup"><span data-stu-id="3a494-557">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="3a494-558">该[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)方法需要**ReadWriteItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="3a494-558">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-559">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-559">Parameters</span></span>

|<span data-ttu-id="3a494-560">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-560">Name</span></span>| <span data-ttu-id="3a494-561">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-561">Type</span></span>| <span data-ttu-id="3a494-562">属性</span><span class="sxs-lookup"><span data-stu-id="3a494-562">Attributes</span></span>| <span data-ttu-id="3a494-563">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-563">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3a494-564">function</span><span class="sxs-lookup"><span data-stu-id="3a494-564">function</span></span>||<span data-ttu-id="3a494-565">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a494-565">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a494-566">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a494-566">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3a494-567">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-567">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="3a494-568">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-568">Object</span></span>| <span data-ttu-id="3a494-569">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-569">&lt;optional&gt;</span></span>|<span data-ttu-id="3a494-570">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a494-570">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3a494-571">错误</span><span class="sxs-lookup"><span data-stu-id="3a494-571">Errors</span></span>

|<span data-ttu-id="3a494-572">错误代码</span><span class="sxs-lookup"><span data-stu-id="3a494-572">Error code</span></span>|<span data-ttu-id="3a494-573">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-573">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3a494-574">请求失败。</span><span class="sxs-lookup"><span data-stu-id="3a494-574">The request has failed.</span></span> <span data-ttu-id="3a494-575">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="3a494-575">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3a494-576">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="3a494-576">The Exchange server returned an error.</span></span> <span data-ttu-id="3a494-577">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-577">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3a494-578">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="3a494-578">The user is no longer connected to the network.</span></span> <span data-ttu-id="3a494-579">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="3a494-579">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-580">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-580">Requirements</span></span>

|<span data-ttu-id="3a494-581">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-581">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="3a494-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-583">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-583">1.0</span></span> | <span data-ttu-id="3a494-584">1.3</span><span class="sxs-lookup"><span data-stu-id="3a494-584">1.3</span></span> |
|[<span data-ttu-id="3a494-585">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-585">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-586">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-586">ReadItem</span></span> | <span data-ttu-id="3a494-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-587">ReadItem</span></span> |
|[<span data-ttu-id="3a494-588">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-589">阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-589">Read</span></span> | <span data-ttu-id="3a494-590">撰写</span><span class="sxs-lookup"><span data-stu-id="3a494-590">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="3a494-591">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-591">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="3a494-592">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a494-592">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3a494-593">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="3a494-593">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="3a494-594">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="3a494-594">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-595">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-595">Parameters</span></span>

|<span data-ttu-id="3a494-596">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-596">Name</span></span>| <span data-ttu-id="3a494-597">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-597">Type</span></span>| <span data-ttu-id="3a494-598">属性</span><span class="sxs-lookup"><span data-stu-id="3a494-598">Attributes</span></span>| <span data-ttu-id="3a494-599">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-599">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3a494-600">function</span><span class="sxs-lookup"><span data-stu-id="3a494-600">function</span></span>||<span data-ttu-id="3a494-601">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a494-601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a494-602">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a494-602">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3a494-603">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-603">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="3a494-604">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-604">Object</span></span>| <span data-ttu-id="3a494-605">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-605">&lt;optional&gt;</span></span>|<span data-ttu-id="3a494-606">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a494-606">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3a494-607">错误</span><span class="sxs-lookup"><span data-stu-id="3a494-607">Errors</span></span>

|<span data-ttu-id="3a494-608">错误代码</span><span class="sxs-lookup"><span data-stu-id="3a494-608">Error code</span></span>|<span data-ttu-id="3a494-609">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-609">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3a494-610">请求失败。</span><span class="sxs-lookup"><span data-stu-id="3a494-610">The request has failed.</span></span> <span data-ttu-id="3a494-611">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="3a494-611">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3a494-612">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="3a494-612">The Exchange server returned an error.</span></span> <span data-ttu-id="3a494-613">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="3a494-613">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3a494-614">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="3a494-614">The user is no longer connected to the network.</span></span> <span data-ttu-id="3a494-615">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="3a494-615">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-616">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-616">Requirements</span></span>

|<span data-ttu-id="3a494-617">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-617">Requirement</span></span>| <span data-ttu-id="3a494-618">值</span><span class="sxs-lookup"><span data-stu-id="3a494-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-619">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-620">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-620">1.0</span></span>|
|[<span data-ttu-id="3a494-621">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-622">ReadItem</span></span>|
|[<span data-ttu-id="3a494-623">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-624">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-624">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-625">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-625">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="3a494-626">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a494-626">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="3a494-627">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="3a494-627">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-628">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="3a494-628">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="3a494-629">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="3a494-629">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="3a494-630">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="3a494-630">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="3a494-631">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="3a494-631">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="3a494-632">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="3a494-632">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="3a494-633">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="3a494-633">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="3a494-634">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="3a494-634">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="3a494-635">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="3a494-635">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="3a494-p150">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="3a494-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="3a494-638">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="3a494-638">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="3a494-639">版本差异</span><span class="sxs-lookup"><span data-stu-id="3a494-639">Version differences</span></span>

<span data-ttu-id="3a494-640">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="3a494-640">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="3a494-641">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。</span><span class="sxs-lookup"><span data-stu-id="3a494-641">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="3a494-642">您可以使用邮箱. hostName 属性确定您的邮件应用程序是在 web 上的 Outlook 中运行还是在桌面客户端上运行。</span><span class="sxs-lookup"><span data-stu-id="3a494-642">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="3a494-643">可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="3a494-643">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-644">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-644">Parameters</span></span>

|<span data-ttu-id="3a494-645">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-645">Name</span></span>| <span data-ttu-id="3a494-646">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-646">Type</span></span>| <span data-ttu-id="3a494-647">属性</span><span class="sxs-lookup"><span data-stu-id="3a494-647">Attributes</span></span>| <span data-ttu-id="3a494-648">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-648">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3a494-649">字符串</span><span class="sxs-lookup"><span data-stu-id="3a494-649">String</span></span>||<span data-ttu-id="3a494-650">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="3a494-650">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="3a494-651">函数</span><span class="sxs-lookup"><span data-stu-id="3a494-651">function</span></span>||<span data-ttu-id="3a494-652">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a494-652">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a494-653">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a494-653">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="3a494-654">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="3a494-654">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="3a494-655">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-655">Object</span></span>| <span data-ttu-id="3a494-656">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-656">&lt;optional&gt;</span></span>|<span data-ttu-id="3a494-657">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a494-657">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-658">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-658">Requirements</span></span>

|<span data-ttu-id="3a494-659">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-659">Requirement</span></span>| <span data-ttu-id="3a494-660">值</span><span class="sxs-lookup"><span data-stu-id="3a494-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-661">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-662">1.0</span><span class="sxs-lookup"><span data-stu-id="3a494-662">1.0</span></span>|
|[<span data-ttu-id="3a494-663">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-664">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3a494-664">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="3a494-665">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-666">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-666">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a494-667">示例</span><span class="sxs-lookup"><span data-stu-id="3a494-667">Example</span></span>

<span data-ttu-id="3a494-668">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="3a494-668">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="3a494-669">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3a494-669">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="3a494-670">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3a494-670">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="3a494-671">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="3a494-671">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a494-672">参数</span><span class="sxs-lookup"><span data-stu-id="3a494-672">Parameters</span></span>

| <span data-ttu-id="3a494-673">名称</span><span class="sxs-lookup"><span data-stu-id="3a494-673">Name</span></span> | <span data-ttu-id="3a494-674">类型</span><span class="sxs-lookup"><span data-stu-id="3a494-674">Type</span></span> | <span data-ttu-id="3a494-675">属性</span><span class="sxs-lookup"><span data-stu-id="3a494-675">Attributes</span></span> | <span data-ttu-id="3a494-676">说明</span><span class="sxs-lookup"><span data-stu-id="3a494-676">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3a494-677">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3a494-677">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3a494-678">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3a494-678">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="3a494-679">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-679">Object</span></span> | <span data-ttu-id="3a494-680">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-680">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-681">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3a494-681">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3a494-682">对象</span><span class="sxs-lookup"><span data-stu-id="3a494-682">Object</span></span> | <span data-ttu-id="3a494-683">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-683">&lt;optional&gt;</span></span> | <span data-ttu-id="3a494-684">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3a494-684">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3a494-685">函数</span><span class="sxs-lookup"><span data-stu-id="3a494-685">function</span></span>| <span data-ttu-id="3a494-686">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a494-686">&lt;optional&gt;</span></span>|<span data-ttu-id="3a494-687">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a494-687">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a494-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a494-688">Requirements</span></span>

|<span data-ttu-id="3a494-689">要求</span><span class="sxs-lookup"><span data-stu-id="3a494-689">Requirement</span></span>| <span data-ttu-id="3a494-690">值</span><span class="sxs-lookup"><span data-stu-id="3a494-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a494-691">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a494-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a494-692">1.5</span><span class="sxs-lookup"><span data-stu-id="3a494-692">1.5</span></span> |
|[<span data-ttu-id="3a494-693">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a494-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a494-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a494-694">ReadItem</span></span> |
|[<span data-ttu-id="3a494-695">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a494-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a494-696">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a494-696">Compose or Read</span></span>|
