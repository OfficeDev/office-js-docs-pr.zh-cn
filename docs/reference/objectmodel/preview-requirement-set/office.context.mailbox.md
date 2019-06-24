---
title: "\"Context\"-\"邮箱-预览要求集\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f2383ea2d2e097b4e2f786bfb1aa8c06ab9eed0e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127595"
---
# <a name="mailbox"></a><span data-ttu-id="3a2d5-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="3a2d5-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="3a2d5-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="3a2d5-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="3a2d5-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a2d5-105">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-105">Requirements</span></span>

|<span data-ttu-id="3a2d5-106">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-106">Requirement</span></span>| <span data-ttu-id="3a2d5-107">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-109">1.0</span></span>|
|[<span data-ttu-id="3a2d5-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-111">受限</span><span class="sxs-lookup"><span data-stu-id="3a2d5-111">Restricted</span></span>|
|[<span data-ttu-id="3a2d5-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3a2d5-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-114">Members and methods</span></span>

| <span data-ttu-id="3a2d5-115">成员</span><span class="sxs-lookup"><span data-stu-id="3a2d5-115">Member</span></span> | <span data-ttu-id="3a2d5-116">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3a2d5-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="3a2d5-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="3a2d5-118">成员</span><span class="sxs-lookup"><span data-stu-id="3a2d5-118">Member</span></span> |
| [<span data-ttu-id="3a2d5-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="3a2d5-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="3a2d5-120">成员</span><span class="sxs-lookup"><span data-stu-id="3a2d5-120">Member</span></span> |
| [<span data-ttu-id="3a2d5-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="3a2d5-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="3a2d5-122">成员</span><span class="sxs-lookup"><span data-stu-id="3a2d5-122">Member</span></span> |
| [<span data-ttu-id="3a2d5-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3a2d5-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3a2d5-124">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-124">Method</span></span> |
| [<span data-ttu-id="3a2d5-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="3a2d5-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="3a2d5-126">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-126">Method</span></span> |
| [<span data-ttu-id="3a2d5-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3a2d5-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="3a2d5-128">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-128">Method</span></span> |
| [<span data-ttu-id="3a2d5-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="3a2d5-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="3a2d5-130">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-130">Method</span></span> |
| [<span data-ttu-id="3a2d5-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="3a2d5-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="3a2d5-132">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-132">Method</span></span> |
| [<span data-ttu-id="3a2d5-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3a2d5-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="3a2d5-134">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-134">Method</span></span> |
| [<span data-ttu-id="3a2d5-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="3a2d5-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="3a2d5-136">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-136">Method</span></span> |
| [<span data-ttu-id="3a2d5-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3a2d5-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="3a2d5-138">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-138">Method</span></span> |
| [<span data-ttu-id="3a2d5-139">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="3a2d5-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="3a2d5-140">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-140">Method</span></span> |
| [<span data-ttu-id="3a2d5-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3a2d5-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="3a2d5-142">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-142">Method</span></span> |
| [<span data-ttu-id="3a2d5-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3a2d5-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="3a2d5-144">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-144">Method</span></span> |
| [<span data-ttu-id="3a2d5-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3a2d5-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="3a2d5-146">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-146">Method</span></span> |
| [<span data-ttu-id="3a2d5-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="3a2d5-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="3a2d5-148">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-148">Method</span></span> |
| [<span data-ttu-id="3a2d5-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3a2d5-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="3a2d5-150">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3a2d5-151">命名空间</span><span class="sxs-lookup"><span data-stu-id="3a2d5-151">Namespaces</span></span>

<span data-ttu-id="3a2d5-152">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="3a2d5-153">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="3a2d5-154">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="3a2d5-155">成员</span><span class="sxs-lookup"><span data-stu-id="3a2d5-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="3a2d5-156">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-156">ewsUrl: String</span></span>

<span data-ttu-id="3a2d5-157">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-157">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="3a2d5-158">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-158">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-159">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a2d5-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3a2d5-162">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="3a2d5-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3a2d5-165">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-165">Type</span></span>

*   <span data-ttu-id="3a2d5-166">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a2d5-167">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-167">Requirements</span></span>

|<span data-ttu-id="3a2d5-168">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-168">Requirement</span></span>| <span data-ttu-id="3a2d5-169">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-171">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-171">1.0</span></span>|
|[<span data-ttu-id="3a2d5-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-173">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-175">Compose or Read</span></span>|

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="3a2d5-176">masterCategories: [masterCategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="3a2d5-177">获取一个对象, 该对象提供用于管理此邮箱上的类别主列表的方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-178">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="3a2d5-179">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-179">Type</span></span>

*   [<span data-ttu-id="3a2d5-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="3a2d5-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="3a2d5-181">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-181">Requirements</span></span>

|<span data-ttu-id="3a2d5-182">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-182">Requirement</span></span>| <span data-ttu-id="3a2d5-183">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-185">预览</span><span class="sxs-lookup"><span data-stu-id="3a2d5-185">Preview</span></span> |
|[<span data-ttu-id="3a2d5-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3a2d5-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="3a2d5-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="3a2d5-190">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-190">Example</span></span>

<span data-ttu-id="3a2d5-191">本示例获取此邮箱的类别主列表。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-191">This example gets the categories master list for this mailbox.</span></span>

```javascript
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="resturl-string"></a><span data-ttu-id="3a2d5-192">Office.context.mailbox.resturl: String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-192">restUrl: String</span></span>

<span data-ttu-id="3a2d5-193">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="3a2d5-194">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="3a2d5-195">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="3a2d5-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3a2d5-198">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-198">Type</span></span>

*   <span data-ttu-id="3a2d5-199">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a2d5-200">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-200">Requirements</span></span>

|<span data-ttu-id="3a2d5-201">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-201">Requirement</span></span>| <span data-ttu-id="3a2d5-202">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-203">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-204">1.5</span><span class="sxs-lookup"><span data-stu-id="3a2d5-204">1.5</span></span> |
|[<span data-ttu-id="3a2d5-205">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-206">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3a2d5-209">方法</span><span class="sxs-lookup"><span data-stu-id="3a2d5-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3a2d5-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3a2d5-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3a2d5-211">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3a2d5-212">目前, 支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-213">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-213">Parameters</span></span>

| <span data-ttu-id="3a2d5-214">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-214">Name</span></span> | <span data-ttu-id="3a2d5-215">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-215">Type</span></span> | <span data-ttu-id="3a2d5-216">属性</span><span class="sxs-lookup"><span data-stu-id="3a2d5-216">Attributes</span></span> | <span data-ttu-id="3a2d5-217">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3a2d5-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3a2d5-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3a2d5-219">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3a2d5-220">函数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-220">Function</span></span> || <span data-ttu-id="3a2d5-p105">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3a2d5-224">Object</span><span class="sxs-lookup"><span data-stu-id="3a2d5-224">Object</span></span> | <span data-ttu-id="3a2d5-225">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-225">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-226">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3a2d5-227">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-227">Object</span></span> | <span data-ttu-id="3a2d5-228">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-228">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-229">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3a2d5-230">函数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-230">function</span></span>| <span data-ttu-id="3a2d5-231">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-231">&lt;optional&gt;</span></span>|<span data-ttu-id="3a2d5-232">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a2d5-233">Requirements</span></span>

|<span data-ttu-id="3a2d5-234">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-234">Requirement</span></span>| <span data-ttu-id="3a2d5-235">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-236">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-237">1.5</span><span class="sxs-lookup"><span data-stu-id="3a2d5-237">1.5</span></span> |
|[<span data-ttu-id="3a2d5-238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-239">ReadItem</span></span> |
|[<span data-ttu-id="3a2d5-240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-242">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-242">Example</span></span>

```javascript
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

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="3a2d5-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3a2d5-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3a2d5-244">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-245">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a2d5-p106">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-248">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-248">Parameters</span></span>

|<span data-ttu-id="3a2d5-249">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-249">Name</span></span>| <span data-ttu-id="3a2d5-250">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-250">Type</span></span>| <span data-ttu-id="3a2d5-251">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a2d5-252">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-252">String</span></span>|<span data-ttu-id="3a2d5-253">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="3a2d5-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3a2d5-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="3a2d5-255">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-256">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-256">Requirements</span></span>

|<span data-ttu-id="3a2d5-257">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-257">Requirement</span></span>| <span data-ttu-id="3a2d5-258">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-260">1.3</span><span class="sxs-lookup"><span data-stu-id="3a2d5-260">1.3</span></span>|
|[<span data-ttu-id="3a2d5-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-262">受限</span><span class="sxs-lookup"><span data-stu-id="3a2d5-262">Restricted</span></span>|
|[<span data-ttu-id="3a2d5-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a2d5-265">返回：</span><span class="sxs-lookup"><span data-stu-id="3a2d5-265">Returns:</span></span>

<span data-ttu-id="3a2d5-266">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3a2d5-267">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-267">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="3a2d5-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="3a2d5-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="3a2d5-269">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="3a2d5-270">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-270">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="3a2d5-271">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-271">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="3a2d5-272">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-272">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="3a2d5-273">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-273">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="3a2d5-274">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-274">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-275">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-275">Parameters</span></span>

|<span data-ttu-id="3a2d5-276">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-276">Name</span></span>| <span data-ttu-id="3a2d5-277">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-277">Type</span></span>| <span data-ttu-id="3a2d5-278">描述</span><span class="sxs-lookup"><span data-stu-id="3a2d5-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="3a2d5-279">日期</span><span class="sxs-lookup"><span data-stu-id="3a2d5-279">Date</span></span>|<span data-ttu-id="3a2d5-280">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-281">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-281">Requirements</span></span>

|<span data-ttu-id="3a2d5-282">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-282">Requirement</span></span>| <span data-ttu-id="3a2d5-283">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-285">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-285">1.0</span></span>|
|[<span data-ttu-id="3a2d5-286">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-287">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-288">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-289">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a2d5-290">返回：</span><span class="sxs-lookup"><span data-stu-id="3a2d5-290">Returns:</span></span>

<span data-ttu-id="3a2d5-291">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="3a2d5-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3a2d5-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3a2d5-293">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-294">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a2d5-p109">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-297">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-297">Parameters</span></span>

|<span data-ttu-id="3a2d5-298">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-298">Name</span></span>| <span data-ttu-id="3a2d5-299">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-299">Type</span></span>| <span data-ttu-id="3a2d5-300">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a2d5-301">字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-301">String</span></span>|<span data-ttu-id="3a2d5-302">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="3a2d5-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3a2d5-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="3a2d5-304">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-305">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-305">Requirements</span></span>

|<span data-ttu-id="3a2d5-306">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-306">Requirement</span></span>| <span data-ttu-id="3a2d5-307">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-309">1.3</span><span class="sxs-lookup"><span data-stu-id="3a2d5-309">1.3</span></span>|
|[<span data-ttu-id="3a2d5-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-311">受限</span><span class="sxs-lookup"><span data-stu-id="3a2d5-311">Restricted</span></span>|
|[<span data-ttu-id="3a2d5-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-313">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a2d5-314">返回：</span><span class="sxs-lookup"><span data-stu-id="3a2d5-314">Returns:</span></span>

<span data-ttu-id="3a2d5-315">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3a2d5-316">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-316">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="3a2d5-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="3a2d5-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="3a2d5-318">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="3a2d5-319">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-320">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-320">Parameters</span></span>

|<span data-ttu-id="3a2d5-321">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-321">Name</span></span>| <span data-ttu-id="3a2d5-322">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-322">Type</span></span>| <span data-ttu-id="3a2d5-323">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="3a2d5-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3a2d5-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="3a2d5-325">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-326">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-326">Requirements</span></span>

|<span data-ttu-id="3a2d5-327">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-327">Requirement</span></span>| <span data-ttu-id="3a2d5-328">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-330">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-330">1.0</span></span>|
|[<span data-ttu-id="3a2d5-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-332">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a2d5-335">返回：</span><span class="sxs-lookup"><span data-stu-id="3a2d5-335">Returns:</span></span>

<span data-ttu-id="3a2d5-336">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-336">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="3a2d5-337">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3a2d5-337">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3a2d5-338">日期</span><span class="sxs-lookup"><span data-stu-id="3a2d5-338">Date</span></span></dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="3a2d5-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="3a2d5-340">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-341">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a2d5-342">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3a2d5-343">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-343">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="3a2d5-344">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-344">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="3a2d5-345">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="3a2d5-346">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-347">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-347">Parameters</span></span>

|<span data-ttu-id="3a2d5-348">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-348">Name</span></span>| <span data-ttu-id="3a2d5-349">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-349">Type</span></span>| <span data-ttu-id="3a2d5-350">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a2d5-351">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-351">String</span></span>|<span data-ttu-id="3a2d5-352">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-353">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-353">Requirements</span></span>

|<span data-ttu-id="3a2d5-354">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-354">Requirement</span></span>| <span data-ttu-id="3a2d5-355">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-356">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-357">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-357">1.0</span></span>|
|[<span data-ttu-id="3a2d5-358">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-359">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-360">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-361">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-362">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-362">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="3a2d5-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="3a2d5-364">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-365">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a2d5-366">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3a2d5-367">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="3a2d5-368">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="3a2d5-p111">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-371">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-371">Parameters</span></span>

|<span data-ttu-id="3a2d5-372">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-372">Name</span></span>| <span data-ttu-id="3a2d5-373">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-373">Type</span></span>| <span data-ttu-id="3a2d5-374">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3a2d5-375">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-375">String</span></span>|<span data-ttu-id="3a2d5-376">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-377">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-377">Requirements</span></span>

|<span data-ttu-id="3a2d5-378">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-378">Requirement</span></span>| <span data-ttu-id="3a2d5-379">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-380">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-381">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-381">1.0</span></span>|
|[<span data-ttu-id="3a2d5-382">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-383">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-384">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-385">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-386">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-386">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="3a2d5-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="3a2d5-388">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-389">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3a2d5-p112">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3a2d5-392">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-392">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="3a2d5-393">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-393">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="3a2d5-394">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-394">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="3a2d5-p114">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="3a2d5-397">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-398">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-399">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-399">All parameters are optional.</span></span>

|<span data-ttu-id="3a2d5-400">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-400">Name</span></span>| <span data-ttu-id="3a2d5-401">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-401">Type</span></span>| <span data-ttu-id="3a2d5-402">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3a2d5-403">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-403">Object</span></span> | <span data-ttu-id="3a2d5-404">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="3a2d5-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a2d5-p115">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="3a2d5-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a2d5-p116">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="3a2d5-411">日期</span><span class="sxs-lookup"><span data-stu-id="3a2d5-411">Date</span></span> | <span data-ttu-id="3a2d5-412">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="3a2d5-413">Date</span><span class="sxs-lookup"><span data-stu-id="3a2d5-413">Date</span></span> | <span data-ttu-id="3a2d5-414">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="3a2d5-415">字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-415">String</span></span> | <span data-ttu-id="3a2d5-p117">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="3a2d5-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="3a2d5-p118">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3a2d5-421">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-421">String</span></span> | <span data-ttu-id="3a2d5-p119">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="3a2d5-424">字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-424">String</span></span> | <span data-ttu-id="3a2d5-p120">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a2d5-427">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-427">Requirements</span></span>

|<span data-ttu-id="3a2d5-428">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-428">Requirement</span></span>| <span data-ttu-id="3a2d5-429">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-431">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-431">1.0</span></span>|
|[<span data-ttu-id="3a2d5-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-433">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-435">阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-436">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-436">Example</span></span>

```javascript
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

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="3a2d5-437">Office.context.mailbox.displaynewmessageform (参数)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="3a2d5-438">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="3a2d5-439">`displayNewMessageForm`方法将打开一个窗体, 使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="3a2d5-440">如果指定了参数, 则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3a2d5-441">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-442">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-443">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-443">All parameters are optional.</span></span>

|<span data-ttu-id="3a2d5-444">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-444">Name</span></span>| <span data-ttu-id="3a2d5-445">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-445">Type</span></span>| <span data-ttu-id="3a2d5-446">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3a2d5-447">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-447">Object</span></span> | <span data-ttu-id="3a2d5-448">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="3a2d5-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a2d5-450">包含电子邮件地址的字符串数组, 或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="3a2d5-451">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="3a2d5-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a2d5-453">包含电子邮件地址的字符串数组, 或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="3a2d5-454">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="3a2d5-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3a2d5-456">包含电子邮件地址的字符串数组, 或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="3a2d5-457">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3a2d5-458">字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-458">String</span></span> | <span data-ttu-id="3a2d5-459">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-459">A string containing the subject of the message.</span></span> <span data-ttu-id="3a2d5-460">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="3a2d5-461">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-461">String</span></span> | <span data-ttu-id="3a2d5-462">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-462">The HTML body of the message.</span></span> <span data-ttu-id="3a2d5-463">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="3a2d5-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3a2d5-465">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="3a2d5-466">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-466">String</span></span> | <span data-ttu-id="3a2d5-p127">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="3a2d5-469">字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-469">String</span></span> | <span data-ttu-id="3a2d5-470">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="3a2d5-471">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-471">String</span></span> | <span data-ttu-id="3a2d5-p128">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="3a2d5-474">布尔</span><span class="sxs-lookup"><span data-stu-id="3a2d5-474">Boolean</span></span> | <span data-ttu-id="3a2d5-p129">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="3a2d5-477">String</span><span class="sxs-lookup"><span data-stu-id="3a2d5-477">String</span></span> | <span data-ttu-id="3a2d5-478">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="3a2d5-479">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="3a2d5-480">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="3a2d5-481">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-481">Requirements</span></span>

|<span data-ttu-id="3a2d5-482">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-482">Requirement</span></span>| <span data-ttu-id="3a2d5-483">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-485">1.6</span><span class="sxs-lookup"><span data-stu-id="3a2d5-485">1.6</span></span> |
|[<span data-ttu-id="3a2d5-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-487">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-489">阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-490">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-490">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="3a2d5-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3a2d5-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="3a2d5-492">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="3a2d5-p131">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-495">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="3a2d5-496">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="3a2d5-496">**REST Tokens**</span></span>

<span data-ttu-id="3a2d5-p132">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="3a2d5-500">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="3a2d5-501">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="3a2d5-501">**EWS Tokens**</span></span>

<span data-ttu-id="3a2d5-p133">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="3a2d5-504">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-505">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-505">Parameters</span></span>

|<span data-ttu-id="3a2d5-506">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-506">Name</span></span>| <span data-ttu-id="3a2d5-507">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-507">Type</span></span>| <span data-ttu-id="3a2d5-508">属性</span><span class="sxs-lookup"><span data-stu-id="3a2d5-508">Attributes</span></span>| <span data-ttu-id="3a2d5-509">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-509">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="3a2d5-510">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-510">Object</span></span> | <span data-ttu-id="3a2d5-511">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-511">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-512">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-512">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="3a2d5-513">布尔值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-513">Boolean</span></span> |  <span data-ttu-id="3a2d5-514">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-514">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-p134">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3a2d5-517">Object</span><span class="sxs-lookup"><span data-stu-id="3a2d5-517">Object</span></span> |  <span data-ttu-id="3a2d5-518">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-518">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-519">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-519">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="3a2d5-520">函数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-520">function</span></span>||<span data-ttu-id="3a2d5-p135">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-523">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-523">Requirements</span></span>

|<span data-ttu-id="3a2d5-524">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-524">Requirement</span></span>| <span data-ttu-id="3a2d5-525">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-526">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-527">1.5</span><span class="sxs-lookup"><span data-stu-id="3a2d5-527">1.5</span></span> |
|[<span data-ttu-id="3a2d5-528">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-528">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-529">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-530">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-530">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-531">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-531">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-532">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-532">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="3a2d5-533">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a2d5-533">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3a2d5-534">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-534">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="3a2d5-p136">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="3a2d5-p137">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3a2d5-540">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-540">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="3a2d5-p138">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-543">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-543">Parameters</span></span>

|<span data-ttu-id="3a2d5-544">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-544">Name</span></span>| <span data-ttu-id="3a2d5-545">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-545">Type</span></span>| <span data-ttu-id="3a2d5-546">属性</span><span class="sxs-lookup"><span data-stu-id="3a2d5-546">Attributes</span></span>| <span data-ttu-id="3a2d5-547">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-547">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3a2d5-548">函数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-548">function</span></span>||<span data-ttu-id="3a2d5-p139">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="3a2d5-551">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-551">Object</span></span>| <span data-ttu-id="3a2d5-552">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-552">&lt;optional&gt;</span></span>|<span data-ttu-id="3a2d5-553">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-553">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-554">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-554">Requirements</span></span>

|<span data-ttu-id="3a2d5-555">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-555">Requirement</span></span>| <span data-ttu-id="3a2d5-556">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-557">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-558">1.3</span><span class="sxs-lookup"><span data-stu-id="3a2d5-558">1.3</span></span>|
|[<span data-ttu-id="3a2d5-559">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-560">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-561">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-562">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-562">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-563">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-563">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="3a2d5-564">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a2d5-564">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3a2d5-565">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-565">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="3a2d5-566">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-566">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-567">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-567">Parameters</span></span>

|<span data-ttu-id="3a2d5-568">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-568">Name</span></span>| <span data-ttu-id="3a2d5-569">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-569">Type</span></span>| <span data-ttu-id="3a2d5-570">属性</span><span class="sxs-lookup"><span data-stu-id="3a2d5-570">Attributes</span></span>| <span data-ttu-id="3a2d5-571">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-571">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3a2d5-572">function</span><span class="sxs-lookup"><span data-stu-id="3a2d5-572">function</span></span>||<span data-ttu-id="3a2d5-573">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-573">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a2d5-574">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-574">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="3a2d5-575">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-575">Object</span></span>| <span data-ttu-id="3a2d5-576">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-576">&lt;optional&gt;</span></span>|<span data-ttu-id="3a2d5-577">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-577">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-578">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-578">Requirements</span></span>

|<span data-ttu-id="3a2d5-579">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-579">Requirement</span></span>| <span data-ttu-id="3a2d5-580">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-581">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-582">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-582">1.0</span></span>|
|[<span data-ttu-id="3a2d5-583">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-583">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-584">ReadItem</span></span>|
|[<span data-ttu-id="3a2d5-585">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-586">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-586">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-587">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-587">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="3a2d5-588">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a2d5-588">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="3a2d5-589">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-589">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-590">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-590">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="3a2d5-591">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="3a2d5-591">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="3a2d5-592">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="3a2d5-592">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="3a2d5-593">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-593">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="3a2d5-594">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-594">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="3a2d5-595">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-595">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="3a2d5-596">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-596">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="3a2d5-597">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-597">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="3a2d5-p141">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="3a2d5-600">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-600">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="3a2d5-601">版本差异</span><span class="sxs-lookup"><span data-stu-id="3a2d5-601">Version differences</span></span>

<span data-ttu-id="3a2d5-602">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-602">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="3a2d5-603">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-603">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="3a2d5-604">您可以使用邮箱. hostName 属性确定您的邮件应用程序是在 web 上的 Outlook 中运行还是在桌面客户端上运行。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-604">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="3a2d5-605">可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-605">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-606">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-606">Parameters</span></span>

|<span data-ttu-id="3a2d5-607">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-607">Name</span></span>| <span data-ttu-id="3a2d5-608">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-608">Type</span></span>| <span data-ttu-id="3a2d5-609">属性</span><span class="sxs-lookup"><span data-stu-id="3a2d5-609">Attributes</span></span>| <span data-ttu-id="3a2d5-610">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-610">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3a2d5-611">字符串</span><span class="sxs-lookup"><span data-stu-id="3a2d5-611">String</span></span>||<span data-ttu-id="3a2d5-612">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-612">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="3a2d5-613">函数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-613">function</span></span>||<span data-ttu-id="3a2d5-614">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a2d5-615">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-615">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="3a2d5-616">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-616">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="3a2d5-617">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-617">Object</span></span>| <span data-ttu-id="3a2d5-618">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-618">&lt;optional&gt;</span></span>|<span data-ttu-id="3a2d5-619">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-619">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-620">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-620">Requirements</span></span>

|<span data-ttu-id="3a2d5-621">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-621">Requirement</span></span>| <span data-ttu-id="3a2d5-622">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-623">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-624">1.0</span><span class="sxs-lookup"><span data-stu-id="3a2d5-624">1.0</span></span>|
|[<span data-ttu-id="3a2d5-625">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-625">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-626">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3a2d5-626">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="3a2d5-627">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-627">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-628">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-628">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a2d5-629">示例</span><span class="sxs-lookup"><span data-stu-id="3a2d5-629">Example</span></span>

<span data-ttu-id="3a2d5-630">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-630">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="3a2d5-631">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3a2d5-631">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="3a2d5-632">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-632">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="3a2d5-633">目前, 支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-633">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a2d5-634">参数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-634">Parameters</span></span>

| <span data-ttu-id="3a2d5-635">名称</span><span class="sxs-lookup"><span data-stu-id="3a2d5-635">Name</span></span> | <span data-ttu-id="3a2d5-636">类型</span><span class="sxs-lookup"><span data-stu-id="3a2d5-636">Type</span></span> | <span data-ttu-id="3a2d5-637">属性</span><span class="sxs-lookup"><span data-stu-id="3a2d5-637">Attributes</span></span> | <span data-ttu-id="3a2d5-638">说明</span><span class="sxs-lookup"><span data-stu-id="3a2d5-638">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3a2d5-639">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3a2d5-639">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3a2d5-640">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-640">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="3a2d5-641">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-641">Object</span></span> | <span data-ttu-id="3a2d5-642">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-642">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-643">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-643">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3a2d5-644">对象</span><span class="sxs-lookup"><span data-stu-id="3a2d5-644">Object</span></span> | <span data-ttu-id="3a2d5-645">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-645">&lt;optional&gt;</span></span> | <span data-ttu-id="3a2d5-646">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-646">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3a2d5-647">函数</span><span class="sxs-lookup"><span data-stu-id="3a2d5-647">function</span></span>| <span data-ttu-id="3a2d5-648">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3a2d5-648">&lt;optional&gt;</span></span>|<span data-ttu-id="3a2d5-649">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3a2d5-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a2d5-650">Requirements</span><span class="sxs-lookup"><span data-stu-id="3a2d5-650">Requirements</span></span>

|<span data-ttu-id="3a2d5-651">要求</span><span class="sxs-lookup"><span data-stu-id="3a2d5-651">Requirement</span></span>| <span data-ttu-id="3a2d5-652">值</span><span class="sxs-lookup"><span data-stu-id="3a2d5-652">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a2d5-653">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3a2d5-653">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a2d5-654">1.5</span><span class="sxs-lookup"><span data-stu-id="3a2d5-654">1.5</span></span> |
|[<span data-ttu-id="3a2d5-655">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3a2d5-655">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a2d5-656">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a2d5-656">ReadItem</span></span> |
|[<span data-ttu-id="3a2d5-657">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3a2d5-657">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3a2d5-658">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3a2d5-658">Compose or Read</span></span>|
