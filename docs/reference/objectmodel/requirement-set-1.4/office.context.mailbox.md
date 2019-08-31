---
title: "\"Context.subname\"-\"邮箱-要求集 1.4\""
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 66ae7cb05ac56224fd7461c5c29587e21a24020a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696209"
---
# <a name="mailbox"></a><span data-ttu-id="701b0-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="701b0-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="701b0-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="701b0-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="701b0-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="701b0-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="701b0-105">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-105">Requirements</span></span>

|<span data-ttu-id="701b0-106">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-106">Requirement</span></span>| <span data-ttu-id="701b0-107">值</span><span class="sxs-lookup"><span data-stu-id="701b0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-109">1.0</span></span>|
|[<span data-ttu-id="701b0-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-111">受限</span><span class="sxs-lookup"><span data-stu-id="701b0-111">Restricted</span></span>|
|[<span data-ttu-id="701b0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="701b0-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="701b0-114">Members and methods</span></span>

| <span data-ttu-id="701b0-115">成员</span><span class="sxs-lookup"><span data-stu-id="701b0-115">Member</span></span> | <span data-ttu-id="701b0-116">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="701b0-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="701b0-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="701b0-118">成员</span><span class="sxs-lookup"><span data-stu-id="701b0-118">Member</span></span> |
| [<span data-ttu-id="701b0-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="701b0-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="701b0-120">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-120">Method</span></span> |
| [<span data-ttu-id="701b0-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="701b0-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="701b0-122">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-122">Method</span></span> |
| [<span data-ttu-id="701b0-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="701b0-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="701b0-124">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-124">Method</span></span> |
| [<span data-ttu-id="701b0-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="701b0-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="701b0-126">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-126">Method</span></span> |
| [<span data-ttu-id="701b0-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="701b0-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="701b0-128">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-128">Method</span></span> |
| [<span data-ttu-id="701b0-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="701b0-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="701b0-130">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-130">Method</span></span> |
| [<span data-ttu-id="701b0-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="701b0-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="701b0-132">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-132">Method</span></span> |
| [<span data-ttu-id="701b0-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="701b0-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="701b0-134">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-134">Method</span></span> |
| [<span data-ttu-id="701b0-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="701b0-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="701b0-136">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-136">Method</span></span> |
| [<span data-ttu-id="701b0-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="701b0-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="701b0-138">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="701b0-139">命名空间</span><span class="sxs-lookup"><span data-stu-id="701b0-139">Namespaces</span></span>

<span data-ttu-id="701b0-140">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="701b0-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="701b0-141">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="701b0-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="701b0-142">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="701b0-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="701b0-143">成员</span><span class="sxs-lookup"><span data-stu-id="701b0-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="701b0-144">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="701b0-144">ewsUrl: String</span></span>

<span data-ttu-id="701b0-145">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="701b0-145">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="701b0-146">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="701b0-146">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-147">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="701b0-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="701b0-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="701b0-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="701b0-150">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="701b0-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="701b0-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="701b0-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="701b0-153">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-153">Type</span></span>

*   <span data-ttu-id="701b0-154">String</span><span class="sxs-lookup"><span data-stu-id="701b0-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="701b0-155">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-155">Requirements</span></span>

|<span data-ttu-id="701b0-156">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-156">Requirement</span></span>| <span data-ttu-id="701b0-157">值</span><span class="sxs-lookup"><span data-stu-id="701b0-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-159">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-159">1.0</span></span>|
|[<span data-ttu-id="701b0-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-161">ReadItem</span></span>|
|[<span data-ttu-id="701b0-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="701b0-164">方法</span><span class="sxs-lookup"><span data-stu-id="701b0-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="701b0-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="701b0-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="701b0-166">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="701b0-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-167">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="701b0-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="701b0-p104">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="701b0-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-170">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-170">Parameters</span></span>

|<span data-ttu-id="701b0-171">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-171">Name</span></span>| <span data-ttu-id="701b0-172">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-172">Type</span></span>| <span data-ttu-id="701b0-173">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="701b0-174">String</span><span class="sxs-lookup"><span data-stu-id="701b0-174">String</span></span>|<span data-ttu-id="701b0-175">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="701b0-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="701b0-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="701b0-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="701b0-177">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="701b0-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-178">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-178">Requirements</span></span>

|<span data-ttu-id="701b0-179">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-179">Requirement</span></span>| <span data-ttu-id="701b0-180">值</span><span class="sxs-lookup"><span data-stu-id="701b0-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-182">1.3</span><span class="sxs-lookup"><span data-stu-id="701b0-182">1.3</span></span>|
|[<span data-ttu-id="701b0-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-184">受限</span><span class="sxs-lookup"><span data-stu-id="701b0-184">Restricted</span></span>|
|[<span data-ttu-id="701b0-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="701b0-187">返回：</span><span class="sxs-lookup"><span data-stu-id="701b0-187">Returns:</span></span>

<span data-ttu-id="701b0-188">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="701b0-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="701b0-189">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="701b0-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="701b0-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="701b0-191">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="701b0-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="701b0-192">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="701b0-192">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="701b0-193">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="701b0-193">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="701b0-194">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="701b0-194">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="701b0-195">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="701b0-195">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="701b0-196">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="701b0-196">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-197">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-197">Parameters</span></span>

|<span data-ttu-id="701b0-198">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-198">Name</span></span>| <span data-ttu-id="701b0-199">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-199">Type</span></span>| <span data-ttu-id="701b0-200">描述</span><span class="sxs-lookup"><span data-stu-id="701b0-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="701b0-201">日期</span><span class="sxs-lookup"><span data-stu-id="701b0-201">Date</span></span>|<span data-ttu-id="701b0-202">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="701b0-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-203">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-203">Requirements</span></span>

|<span data-ttu-id="701b0-204">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-204">Requirement</span></span>| <span data-ttu-id="701b0-205">值</span><span class="sxs-lookup"><span data-stu-id="701b0-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-207">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-207">1.0</span></span>|
|[<span data-ttu-id="701b0-208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-209">ReadItem</span></span>|
|[<span data-ttu-id="701b0-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="701b0-212">返回：</span><span class="sxs-lookup"><span data-stu-id="701b0-212">Returns:</span></span>

<span data-ttu-id="701b0-213">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="701b0-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="701b0-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="701b0-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="701b0-215">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="701b0-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-216">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="701b0-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="701b0-p107">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="701b0-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-219">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-219">Parameters</span></span>

|<span data-ttu-id="701b0-220">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-220">Name</span></span>| <span data-ttu-id="701b0-221">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-221">Type</span></span>| <span data-ttu-id="701b0-222">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="701b0-223">字符串</span><span class="sxs-lookup"><span data-stu-id="701b0-223">String</span></span>|<span data-ttu-id="701b0-224">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="701b0-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="701b0-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="701b0-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="701b0-226">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="701b0-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-227">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-227">Requirements</span></span>

|<span data-ttu-id="701b0-228">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-228">Requirement</span></span>| <span data-ttu-id="701b0-229">值</span><span class="sxs-lookup"><span data-stu-id="701b0-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-230">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-231">1.3</span><span class="sxs-lookup"><span data-stu-id="701b0-231">1.3</span></span>|
|[<span data-ttu-id="701b0-232">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-233">受限</span><span class="sxs-lookup"><span data-stu-id="701b0-233">Restricted</span></span>|
|[<span data-ttu-id="701b0-234">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-235">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="701b0-236">返回：</span><span class="sxs-lookup"><span data-stu-id="701b0-236">Returns:</span></span>

<span data-ttu-id="701b0-237">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="701b0-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="701b0-238">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="701b0-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="701b0-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="701b0-240">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="701b0-241">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-242">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-242">Parameters</span></span>

|<span data-ttu-id="701b0-243">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-243">Name</span></span>| <span data-ttu-id="701b0-244">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-244">Type</span></span>| <span data-ttu-id="701b0-245">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="701b0-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="701b0-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|<span data-ttu-id="701b0-247">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="701b0-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-248">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-248">Requirements</span></span>

|<span data-ttu-id="701b0-249">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-249">Requirement</span></span>| <span data-ttu-id="701b0-250">值</span><span class="sxs-lookup"><span data-stu-id="701b0-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-252">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-252">1.0</span></span>|
|[<span data-ttu-id="701b0-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-254">ReadItem</span></span>|
|[<span data-ttu-id="701b0-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-256">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="701b0-257">返回：</span><span class="sxs-lookup"><span data-stu-id="701b0-257">Returns:</span></span>

<span data-ttu-id="701b0-258">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="701b0-259">类型: Date</span><span class="sxs-lookup"><span data-stu-id="701b0-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="701b0-260">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="701b0-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="701b0-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="701b0-262">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="701b0-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-263">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="701b0-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="701b0-264">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="701b0-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="701b0-265">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="701b0-265">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="701b0-266">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="701b0-266">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="701b0-267">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="701b0-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="701b0-268">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="701b0-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-269">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-269">Parameters</span></span>

|<span data-ttu-id="701b0-270">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-270">Name</span></span>| <span data-ttu-id="701b0-271">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-271">Type</span></span>| <span data-ttu-id="701b0-272">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="701b0-273">String</span><span class="sxs-lookup"><span data-stu-id="701b0-273">String</span></span>|<span data-ttu-id="701b0-274">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="701b0-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-275">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-275">Requirements</span></span>

|<span data-ttu-id="701b0-276">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-276">Requirement</span></span>| <span data-ttu-id="701b0-277">值</span><span class="sxs-lookup"><span data-stu-id="701b0-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-279">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-279">1.0</span></span>|
|[<span data-ttu-id="701b0-280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-281">ReadItem</span></span>|
|[<span data-ttu-id="701b0-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-283">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="701b0-284">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="701b0-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="701b0-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="701b0-286">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="701b0-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-287">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="701b0-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="701b0-288">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="701b0-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="701b0-289">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="701b0-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="701b0-290">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="701b0-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="701b0-p109">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="701b0-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-293">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-293">Parameters</span></span>

|<span data-ttu-id="701b0-294">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-294">Name</span></span>| <span data-ttu-id="701b0-295">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-295">Type</span></span>| <span data-ttu-id="701b0-296">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="701b0-297">String</span><span class="sxs-lookup"><span data-stu-id="701b0-297">String</span></span>|<span data-ttu-id="701b0-298">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="701b0-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-299">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-299">Requirements</span></span>

|<span data-ttu-id="701b0-300">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-300">Requirement</span></span>| <span data-ttu-id="701b0-301">值</span><span class="sxs-lookup"><span data-stu-id="701b0-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-303">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-303">1.0</span></span>|
|[<span data-ttu-id="701b0-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-305">ReadItem</span></span>|
|[<span data-ttu-id="701b0-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-307">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="701b0-308">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="701b0-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="701b0-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="701b0-310">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="701b0-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-311">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="701b0-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="701b0-p110">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="701b0-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="701b0-314">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="701b0-314">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="701b0-315">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="701b0-315">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="701b0-316">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="701b0-316">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="701b0-p112">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="701b0-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="701b0-319">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="701b0-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-320">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-320">Parameters</span></span>

|<span data-ttu-id="701b0-321">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-321">Name</span></span>| <span data-ttu-id="701b0-322">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-322">Type</span></span>| <span data-ttu-id="701b0-323">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="701b0-324">对象</span><span class="sxs-lookup"><span data-stu-id="701b0-324">Object</span></span> | <span data-ttu-id="701b0-325">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="701b0-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="701b0-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="701b0-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="701b0-p113">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="701b0-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="701b0-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="701b0-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="701b0-p114">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="701b0-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="701b0-332">日期</span><span class="sxs-lookup"><span data-stu-id="701b0-332">Date</span></span> | <span data-ttu-id="701b0-333">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="701b0-334">Date</span><span class="sxs-lookup"><span data-stu-id="701b0-334">Date</span></span> | <span data-ttu-id="701b0-335">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="701b0-336">String</span><span class="sxs-lookup"><span data-stu-id="701b0-336">String</span></span> | <span data-ttu-id="701b0-p115">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="701b0-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="701b0-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="701b0-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="701b0-p116">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="701b0-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="701b0-342">String</span><span class="sxs-lookup"><span data-stu-id="701b0-342">String</span></span> | <span data-ttu-id="701b0-p117">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="701b0-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="701b0-345">字符串</span><span class="sxs-lookup"><span data-stu-id="701b0-345">String</span></span> | <span data-ttu-id="701b0-p118">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="701b0-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="701b0-348">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-348">Requirements</span></span>

|<span data-ttu-id="701b0-349">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-349">Requirement</span></span>| <span data-ttu-id="701b0-350">值</span><span class="sxs-lookup"><span data-stu-id="701b0-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-351">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-352">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-352">1.0</span></span>|
|[<span data-ttu-id="701b0-353">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-354">ReadItem</span></span>|
|[<span data-ttu-id="701b0-355">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-356">阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="701b0-357">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="701b0-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="701b0-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="701b0-359">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="701b0-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="701b0-p119">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="701b0-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="701b0-p120">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="701b0-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="701b0-365">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="701b0-365">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="701b0-p121">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="701b0-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-368">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-368">Parameters</span></span>

|<span data-ttu-id="701b0-369">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-369">Name</span></span>| <span data-ttu-id="701b0-370">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-370">Type</span></span>| <span data-ttu-id="701b0-371">属性</span><span class="sxs-lookup"><span data-stu-id="701b0-371">Attributes</span></span>| <span data-ttu-id="701b0-372">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="701b0-373">function</span><span class="sxs-lookup"><span data-stu-id="701b0-373">function</span></span>||<span data-ttu-id="701b0-374">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="701b0-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="701b0-375">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="701b0-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="701b0-376">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="701b0-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="701b0-377">对象</span><span class="sxs-lookup"><span data-stu-id="701b0-377">Object</span></span>| <span data-ttu-id="701b0-378">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="701b0-378">&lt;optional&gt;</span></span>|<span data-ttu-id="701b0-379">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="701b0-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="701b0-380">错误</span><span class="sxs-lookup"><span data-stu-id="701b0-380">Errors</span></span>

|<span data-ttu-id="701b0-381">错误代码</span><span class="sxs-lookup"><span data-stu-id="701b0-381">Error code</span></span>|<span data-ttu-id="701b0-382">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="701b0-383">请求失败。</span><span class="sxs-lookup"><span data-stu-id="701b0-383">The request has failed.</span></span> <span data-ttu-id="701b0-384">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="701b0-385">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="701b0-385">The Exchange server returned an error.</span></span> <span data-ttu-id="701b0-386">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="701b0-387">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="701b0-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="701b0-388">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="701b0-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-389">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-389">Requirements</span></span>

|<span data-ttu-id="701b0-390">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-390">Requirement</span></span>| <span data-ttu-id="701b0-391">值</span><span class="sxs-lookup"><span data-stu-id="701b0-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-392">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-393">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-393">1.0</span></span>|
|[<span data-ttu-id="701b0-394">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-395">ReadItem</span></span>|
|[<span data-ttu-id="701b0-396">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-397">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-397">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="701b0-398">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-398">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="701b0-399">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="701b0-399">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="701b0-400">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="701b0-400">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="701b0-401">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="701b0-401">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-402">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-402">Parameters</span></span>

|<span data-ttu-id="701b0-403">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-403">Name</span></span>| <span data-ttu-id="701b0-404">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-404">Type</span></span>| <span data-ttu-id="701b0-405">属性</span><span class="sxs-lookup"><span data-stu-id="701b0-405">Attributes</span></span>| <span data-ttu-id="701b0-406">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-406">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="701b0-407">function</span><span class="sxs-lookup"><span data-stu-id="701b0-407">function</span></span>||<span data-ttu-id="701b0-408">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="701b0-408">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="701b0-409">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="701b0-409">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="701b0-410">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="701b0-410">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="701b0-411">对象</span><span class="sxs-lookup"><span data-stu-id="701b0-411">Object</span></span>| <span data-ttu-id="701b0-412">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="701b0-412">&lt;optional&gt;</span></span>|<span data-ttu-id="701b0-413">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="701b0-413">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="701b0-414">错误</span><span class="sxs-lookup"><span data-stu-id="701b0-414">Errors</span></span>

|<span data-ttu-id="701b0-415">错误代码</span><span class="sxs-lookup"><span data-stu-id="701b0-415">Error code</span></span>|<span data-ttu-id="701b0-416">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-416">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="701b0-417">请求失败。</span><span class="sxs-lookup"><span data-stu-id="701b0-417">The request has failed.</span></span> <span data-ttu-id="701b0-418">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-418">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="701b0-419">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="701b0-419">The Exchange server returned an error.</span></span> <span data-ttu-id="701b0-420">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="701b0-420">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="701b0-421">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="701b0-421">The user is no longer connected to the network.</span></span> <span data-ttu-id="701b0-422">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="701b0-422">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-423">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-423">Requirements</span></span>

|<span data-ttu-id="701b0-424">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-424">Requirement</span></span>| <span data-ttu-id="701b0-425">值</span><span class="sxs-lookup"><span data-stu-id="701b0-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-426">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-427">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-427">1.0</span></span>|
|[<span data-ttu-id="701b0-428">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="701b0-429">ReadItem</span></span>|
|[<span data-ttu-id="701b0-430">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-431">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="701b0-432">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-432">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="701b0-433">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="701b0-433">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="701b0-434">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="701b0-434">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-435">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="701b0-435">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="701b0-436">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="701b0-436">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="701b0-437">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="701b0-437">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="701b0-438">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="701b0-438">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="701b0-439">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="701b0-439">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="701b0-440">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="701b0-440">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="701b0-441">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="701b0-441">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="701b0-442">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="701b0-442">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="701b0-p129">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="701b0-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="701b0-445">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="701b0-445">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="701b0-446">版本差异</span><span class="sxs-lookup"><span data-stu-id="701b0-446">Version differences</span></span>

<span data-ttu-id="701b0-447">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="701b0-447">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="701b0-p130">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="701b0-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="701b0-451">参数</span><span class="sxs-lookup"><span data-stu-id="701b0-451">Parameters</span></span>

|<span data-ttu-id="701b0-452">名称</span><span class="sxs-lookup"><span data-stu-id="701b0-452">Name</span></span>| <span data-ttu-id="701b0-453">类型</span><span class="sxs-lookup"><span data-stu-id="701b0-453">Type</span></span>| <span data-ttu-id="701b0-454">属性</span><span class="sxs-lookup"><span data-stu-id="701b0-454">Attributes</span></span>| <span data-ttu-id="701b0-455">说明</span><span class="sxs-lookup"><span data-stu-id="701b0-455">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="701b0-456">字符串</span><span class="sxs-lookup"><span data-stu-id="701b0-456">String</span></span>||<span data-ttu-id="701b0-457">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="701b0-457">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="701b0-458">函数</span><span class="sxs-lookup"><span data-stu-id="701b0-458">function</span></span>||<span data-ttu-id="701b0-459">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="701b0-459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="701b0-460">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="701b0-460">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="701b0-461">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="701b0-461">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="701b0-462">对象</span><span class="sxs-lookup"><span data-stu-id="701b0-462">Object</span></span>| <span data-ttu-id="701b0-463">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="701b0-463">&lt;optional&gt;</span></span>|<span data-ttu-id="701b0-464">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="701b0-464">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="701b0-465">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-465">Requirements</span></span>

|<span data-ttu-id="701b0-466">要求</span><span class="sxs-lookup"><span data-stu-id="701b0-466">Requirement</span></span>| <span data-ttu-id="701b0-467">值</span><span class="sxs-lookup"><span data-stu-id="701b0-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="701b0-468">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="701b0-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="701b0-469">1.0</span><span class="sxs-lookup"><span data-stu-id="701b0-469">1.0</span></span>|
|[<span data-ttu-id="701b0-470">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="701b0-470">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="701b0-471">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="701b0-471">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="701b0-472">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="701b0-472">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="701b0-473">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="701b0-473">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="701b0-474">示例</span><span class="sxs-lookup"><span data-stu-id="701b0-474">Example</span></span>

<span data-ttu-id="701b0-475">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="701b0-475">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
