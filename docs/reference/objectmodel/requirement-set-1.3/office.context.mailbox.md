---
title: "\"Context.subname\"-\"邮箱-要求集 1.3\""
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: f1896803c38abd03f63b0a9ae689d91eeb5540de
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627018"
---
# <a name="mailbox"></a><span data-ttu-id="cf0bd-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="cf0bd-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="cf0bd-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="cf0bd-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="cf0bd-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf0bd-105">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-105">Requirements</span></span>

|<span data-ttu-id="cf0bd-106">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-106">Requirement</span></span>| <span data-ttu-id="cf0bd-107">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-109">1.0</span></span>|
|[<span data-ttu-id="cf0bd-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-111">受限</span><span class="sxs-lookup"><span data-stu-id="cf0bd-111">Restricted</span></span>|
|[<span data-ttu-id="cf0bd-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cf0bd-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-114">Members and methods</span></span>

| <span data-ttu-id="cf0bd-115">成员</span><span class="sxs-lookup"><span data-stu-id="cf0bd-115">Member</span></span> | <span data-ttu-id="cf0bd-116">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cf0bd-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="cf0bd-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="cf0bd-118">成员</span><span class="sxs-lookup"><span data-stu-id="cf0bd-118">Member</span></span> |
| [<span data-ttu-id="cf0bd-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="cf0bd-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="cf0bd-120">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-120">Method</span></span> |
| [<span data-ttu-id="cf0bd-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="cf0bd-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="cf0bd-122">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-122">Method</span></span> |
| [<span data-ttu-id="cf0bd-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="cf0bd-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="cf0bd-124">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-124">Method</span></span> |
| [<span data-ttu-id="cf0bd-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="cf0bd-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="cf0bd-126">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-126">Method</span></span> |
| [<span data-ttu-id="cf0bd-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="cf0bd-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="cf0bd-128">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-128">Method</span></span> |
| [<span data-ttu-id="cf0bd-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="cf0bd-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="cf0bd-130">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-130">Method</span></span> |
| [<span data-ttu-id="cf0bd-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="cf0bd-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="cf0bd-132">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-132">Method</span></span> |
| [<span data-ttu-id="cf0bd-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="cf0bd-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="cf0bd-134">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-134">Method</span></span> |
| [<span data-ttu-id="cf0bd-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="cf0bd-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="cf0bd-136">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-136">Method</span></span> |
| [<span data-ttu-id="cf0bd-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="cf0bd-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="cf0bd-138">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="cf0bd-139">命名空间</span><span class="sxs-lookup"><span data-stu-id="cf0bd-139">Namespaces</span></span>

<span data-ttu-id="cf0bd-140">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="cf0bd-141">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="cf0bd-142">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="cf0bd-143">Members</span><span class="sxs-lookup"><span data-stu-id="cf0bd-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="cf0bd-144">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-144">ewsUrl: String</span></span>

<span data-ttu-id="cf0bd-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-147">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf0bd-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="cf0bd-150">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="cf0bd-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="cf0bd-153">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-153">Type</span></span>

*   <span data-ttu-id="cf0bd-154">String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf0bd-155">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-155">Requirements</span></span>

|<span data-ttu-id="cf0bd-156">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-156">Requirement</span></span>| <span data-ttu-id="cf0bd-157">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-159">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-159">1.0</span></span>|
|[<span data-ttu-id="cf0bd-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-161">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="cf0bd-164">方法</span><span class="sxs-lookup"><span data-stu-id="cf0bd-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="cf0bd-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="cf0bd-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="cf0bd-166">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-167">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf0bd-p104">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-170">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-170">Parameters</span></span>

|<span data-ttu-id="cf0bd-171">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-171">Name</span></span>| <span data-ttu-id="cf0bd-172">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-172">Type</span></span>| <span data-ttu-id="cf0bd-173">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cf0bd-174">String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-174">String</span></span>|<span data-ttu-id="cf0bd-175">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="cf0bd-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="cf0bd-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="cf0bd-177">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-178">Requirements</span></span>

|<span data-ttu-id="cf0bd-179">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-179">Requirement</span></span>| <span data-ttu-id="cf0bd-180">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-182">1.3</span><span class="sxs-lookup"><span data-stu-id="cf0bd-182">1.3</span></span>|
|[<span data-ttu-id="cf0bd-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-184">受限</span><span class="sxs-lookup"><span data-stu-id="cf0bd-184">Restricted</span></span>|
|[<span data-ttu-id="cf0bd-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf0bd-187">返回：</span><span class="sxs-lookup"><span data-stu-id="cf0bd-187">Returns:</span></span>

<span data-ttu-id="cf0bd-188">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="cf0bd-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="cf0bd-189">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="cf0bd-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="cf0bd-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="cf0bd-191">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="cf0bd-p105">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p105">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="cf0bd-p106">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p106">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-197">Parameters</span><span class="sxs-lookup"><span data-stu-id="cf0bd-197">Parameters</span></span>

|<span data-ttu-id="cf0bd-198">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-198">Name</span></span>| <span data-ttu-id="cf0bd-199">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-199">Type</span></span>| <span data-ttu-id="cf0bd-200">描述</span><span class="sxs-lookup"><span data-stu-id="cf0bd-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="cf0bd-201">日期</span><span class="sxs-lookup"><span data-stu-id="cf0bd-201">Date</span></span>|<span data-ttu-id="cf0bd-202">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="cf0bd-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-203">Requirements</span></span>

|<span data-ttu-id="cf0bd-204">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-204">Requirement</span></span>| <span data-ttu-id="cf0bd-205">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-207">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-207">1.0</span></span>|
|[<span data-ttu-id="cf0bd-208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-209">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf0bd-212">返回：</span><span class="sxs-lookup"><span data-stu-id="cf0bd-212">Returns:</span></span>

<span data-ttu-id="cf0bd-213">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="cf0bd-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="cf0bd-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="cf0bd-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="cf0bd-215">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-216">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf0bd-p107">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-219">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-219">Parameters</span></span>

|<span data-ttu-id="cf0bd-220">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-220">Name</span></span>| <span data-ttu-id="cf0bd-221">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-221">Type</span></span>| <span data-ttu-id="cf0bd-222">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cf0bd-223">字符串</span><span class="sxs-lookup"><span data-stu-id="cf0bd-223">String</span></span>|<span data-ttu-id="cf0bd-224">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="cf0bd-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="cf0bd-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="cf0bd-226">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-227">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-227">Requirements</span></span>

|<span data-ttu-id="cf0bd-228">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-228">Requirement</span></span>| <span data-ttu-id="cf0bd-229">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-230">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-231">1.3</span><span class="sxs-lookup"><span data-stu-id="cf0bd-231">1.3</span></span>|
|[<span data-ttu-id="cf0bd-232">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-233">受限</span><span class="sxs-lookup"><span data-stu-id="cf0bd-233">Restricted</span></span>|
|[<span data-ttu-id="cf0bd-234">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-235">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf0bd-236">返回：</span><span class="sxs-lookup"><span data-stu-id="cf0bd-236">Returns:</span></span>

<span data-ttu-id="cf0bd-237">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="cf0bd-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="cf0bd-238">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="cf0bd-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="cf0bd-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="cf0bd-240">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="cf0bd-241">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-242">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-242">Parameters</span></span>

|<span data-ttu-id="cf0bd-243">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-243">Name</span></span>| <span data-ttu-id="cf0bd-244">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-244">Type</span></span>| <span data-ttu-id="cf0bd-245">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="cf0bd-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="cf0bd-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="cf0bd-247">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-248">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-248">Requirements</span></span>

|<span data-ttu-id="cf0bd-249">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-249">Requirement</span></span>| <span data-ttu-id="cf0bd-250">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-252">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-252">1.0</span></span>|
|[<span data-ttu-id="cf0bd-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-254">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-256">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf0bd-257">返回：</span><span class="sxs-lookup"><span data-stu-id="cf0bd-257">Returns:</span></span>

<span data-ttu-id="cf0bd-258">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="cf0bd-259">键入：日期</span><span class="sxs-lookup"><span data-stu-id="cf0bd-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="cf0bd-260">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="cf0bd-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="cf0bd-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="cf0bd-262">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-263">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf0bd-264">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="cf0bd-p108">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p108">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="cf0bd-267">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="cf0bd-268">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-269">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-269">Parameters</span></span>

|<span data-ttu-id="cf0bd-270">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-270">Name</span></span>| <span data-ttu-id="cf0bd-271">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-271">Type</span></span>| <span data-ttu-id="cf0bd-272">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cf0bd-273">String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-273">String</span></span>|<span data-ttu-id="cf0bd-274">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-275">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-275">Requirements</span></span>

|<span data-ttu-id="cf0bd-276">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-276">Requirement</span></span>| <span data-ttu-id="cf0bd-277">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-279">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-279">1.0</span></span>|
|[<span data-ttu-id="cf0bd-280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-281">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-283">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf0bd-284">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="cf0bd-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="cf0bd-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="cf0bd-286">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-287">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf0bd-288">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="cf0bd-289">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="cf0bd-290">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="cf0bd-p109">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-293">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-293">Parameters</span></span>

|<span data-ttu-id="cf0bd-294">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-294">Name</span></span>| <span data-ttu-id="cf0bd-295">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-295">Type</span></span>| <span data-ttu-id="cf0bd-296">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cf0bd-297">String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-297">String</span></span>|<span data-ttu-id="cf0bd-298">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-299">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-299">Requirements</span></span>

|<span data-ttu-id="cf0bd-300">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-300">Requirement</span></span>| <span data-ttu-id="cf0bd-301">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-303">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-303">1.0</span></span>|
|[<span data-ttu-id="cf0bd-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-305">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-307">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf0bd-308">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="cf0bd-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="cf0bd-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="cf0bd-310">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-311">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf0bd-p110">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="cf0bd-p111">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p111">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="cf0bd-p112">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="cf0bd-319">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-320">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-320">Parameters</span></span>

|<span data-ttu-id="cf0bd-321">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-321">Name</span></span>| <span data-ttu-id="cf0bd-322">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-322">Type</span></span>| <span data-ttu-id="cf0bd-323">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="cf0bd-324">对象</span><span class="sxs-lookup"><span data-stu-id="cf0bd-324">Object</span></span> | <span data-ttu-id="cf0bd-325">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="cf0bd-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="cf0bd-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="cf0bd-p113">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="cf0bd-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="cf0bd-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="cf0bd-p114">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="cf0bd-332">日期</span><span class="sxs-lookup"><span data-stu-id="cf0bd-332">Date</span></span> | <span data-ttu-id="cf0bd-333">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="cf0bd-334">Date</span><span class="sxs-lookup"><span data-stu-id="cf0bd-334">Date</span></span> | <span data-ttu-id="cf0bd-335">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="cf0bd-336">String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-336">String</span></span> | <span data-ttu-id="cf0bd-p115">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="cf0bd-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="cf0bd-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="cf0bd-p116">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="cf0bd-342">String</span><span class="sxs-lookup"><span data-stu-id="cf0bd-342">String</span></span> | <span data-ttu-id="cf0bd-p117">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="cf0bd-345">字符串</span><span class="sxs-lookup"><span data-stu-id="cf0bd-345">String</span></span> | <span data-ttu-id="cf0bd-p118">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf0bd-348">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-348">Requirements</span></span>

|<span data-ttu-id="cf0bd-349">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-349">Requirement</span></span>| <span data-ttu-id="cf0bd-350">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-351">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-352">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-352">1.0</span></span>|
|[<span data-ttu-id="cf0bd-353">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-354">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-355">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-356">阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf0bd-357">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="cf0bd-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cf0bd-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="cf0bd-359">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="cf0bd-p119">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="cf0bd-362">您可以将令牌和附件标识符或项目标识符同时传递给第三方系统。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-362">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="cf0bd-363">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-363">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="cf0bd-364">例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-364">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="cf0bd-365">在读取`getCallbackTokenAsync`模式下调用方法需要**ReadItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-365">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="cf0bd-366">在`getCallbackTokenAsync`撰写模式下调用需要您保存项目。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-366">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="cf0bd-367">该[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)方法需要**ReadWriteItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-367">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-368">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-368">Parameters</span></span>

|<span data-ttu-id="cf0bd-369">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-369">Name</span></span>| <span data-ttu-id="cf0bd-370">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-370">Type</span></span>| <span data-ttu-id="cf0bd-371">属性</span><span class="sxs-lookup"><span data-stu-id="cf0bd-371">Attributes</span></span>| <span data-ttu-id="cf0bd-372">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cf0bd-373">function</span><span class="sxs-lookup"><span data-stu-id="cf0bd-373">function</span></span>||<span data-ttu-id="cf0bd-374">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cf0bd-375">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="cf0bd-376">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="cf0bd-377">对象</span><span class="sxs-lookup"><span data-stu-id="cf0bd-377">Object</span></span>| <span data-ttu-id="cf0bd-378">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cf0bd-378">&lt;optional&gt;</span></span>|<span data-ttu-id="cf0bd-379">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cf0bd-380">错误</span><span class="sxs-lookup"><span data-stu-id="cf0bd-380">Errors</span></span>

|<span data-ttu-id="cf0bd-381">错误代码</span><span class="sxs-lookup"><span data-stu-id="cf0bd-381">Error code</span></span>|<span data-ttu-id="cf0bd-382">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="cf0bd-383">请求失败。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-383">The request has failed.</span></span> <span data-ttu-id="cf0bd-384">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="cf0bd-385">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-385">The Exchange server returned an error.</span></span> <span data-ttu-id="cf0bd-386">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="cf0bd-387">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="cf0bd-388">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-389">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-389">Requirements</span></span>

|<span data-ttu-id="cf0bd-390">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="cf0bd-391">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-392">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-392">1.0</span></span> | <span data-ttu-id="cf0bd-393">1.3</span><span class="sxs-lookup"><span data-stu-id="cf0bd-393">1.3</span></span> |
|[<span data-ttu-id="cf0bd-394">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-395">ReadItem</span></span> | <span data-ttu-id="cf0bd-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-396">ReadItem</span></span> |
|[<span data-ttu-id="cf0bd-397">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-398">阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-398">Read</span></span> | <span data-ttu-id="cf0bd-399">撰写</span><span class="sxs-lookup"><span data-stu-id="cf0bd-399">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="cf0bd-400">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-400">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="cf0bd-401">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cf0bd-401">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="cf0bd-402">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-402">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="cf0bd-403">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-403">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-404">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-404">Parameters</span></span>

|<span data-ttu-id="cf0bd-405">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-405">Name</span></span>| <span data-ttu-id="cf0bd-406">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-406">Type</span></span>| <span data-ttu-id="cf0bd-407">属性</span><span class="sxs-lookup"><span data-stu-id="cf0bd-407">Attributes</span></span>| <span data-ttu-id="cf0bd-408">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-408">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cf0bd-409">function</span><span class="sxs-lookup"><span data-stu-id="cf0bd-409">function</span></span>||<span data-ttu-id="cf0bd-410">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-410">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cf0bd-411">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-411">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="cf0bd-412">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-412">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="cf0bd-413">对象</span><span class="sxs-lookup"><span data-stu-id="cf0bd-413">Object</span></span>| <span data-ttu-id="cf0bd-414">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cf0bd-414">&lt;optional&gt;</span></span>|<span data-ttu-id="cf0bd-415">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-415">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cf0bd-416">错误</span><span class="sxs-lookup"><span data-stu-id="cf0bd-416">Errors</span></span>

|<span data-ttu-id="cf0bd-417">错误代码</span><span class="sxs-lookup"><span data-stu-id="cf0bd-417">Error code</span></span>|<span data-ttu-id="cf0bd-418">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-418">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="cf0bd-419">请求失败。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-419">The request has failed.</span></span> <span data-ttu-id="cf0bd-420">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-420">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="cf0bd-421">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-421">The Exchange server returned an error.</span></span> <span data-ttu-id="cf0bd-422">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-422">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="cf0bd-423">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-423">The user is no longer connected to the network.</span></span> <span data-ttu-id="cf0bd-424">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-424">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-425">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-425">Requirements</span></span>

|<span data-ttu-id="cf0bd-426">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-426">Requirement</span></span>| <span data-ttu-id="cf0bd-427">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-428">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-429">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-429">1.0</span></span>|
|[<span data-ttu-id="cf0bd-430">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf0bd-431">ReadItem</span></span>|
|[<span data-ttu-id="cf0bd-432">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-433">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf0bd-434">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-434">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="cf0bd-435">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cf0bd-435">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="cf0bd-436">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-436">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-437">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-437">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="cf0bd-438">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="cf0bd-438">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="cf0bd-439">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="cf0bd-439">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="cf0bd-440">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-440">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="cf0bd-441">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-441">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="cf0bd-442">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-442">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="cf0bd-443">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-443">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="cf0bd-444">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-444">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="cf0bd-p129">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="cf0bd-447">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-447">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="cf0bd-448">版本差异</span><span class="sxs-lookup"><span data-stu-id="cf0bd-448">Version differences</span></span>

<span data-ttu-id="cf0bd-449">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-449">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="cf0bd-p130">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf0bd-453">参数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-453">Parameters</span></span>

|<span data-ttu-id="cf0bd-454">名称</span><span class="sxs-lookup"><span data-stu-id="cf0bd-454">Name</span></span>| <span data-ttu-id="cf0bd-455">类型</span><span class="sxs-lookup"><span data-stu-id="cf0bd-455">Type</span></span>| <span data-ttu-id="cf0bd-456">属性</span><span class="sxs-lookup"><span data-stu-id="cf0bd-456">Attributes</span></span>| <span data-ttu-id="cf0bd-457">说明</span><span class="sxs-lookup"><span data-stu-id="cf0bd-457">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="cf0bd-458">字符串</span><span class="sxs-lookup"><span data-stu-id="cf0bd-458">String</span></span>||<span data-ttu-id="cf0bd-459">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-459">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="cf0bd-460">函数</span><span class="sxs-lookup"><span data-stu-id="cf0bd-460">function</span></span>||<span data-ttu-id="cf0bd-461">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-461">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cf0bd-462">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-462">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="cf0bd-463">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-463">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="cf0bd-464">对象</span><span class="sxs-lookup"><span data-stu-id="cf0bd-464">Object</span></span>| <span data-ttu-id="cf0bd-465">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cf0bd-465">&lt;optional&gt;</span></span>|<span data-ttu-id="cf0bd-466">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-466">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf0bd-467">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf0bd-467">Requirements</span></span>

|<span data-ttu-id="cf0bd-468">要求</span><span class="sxs-lookup"><span data-stu-id="cf0bd-468">Requirement</span></span>| <span data-ttu-id="cf0bd-469">值</span><span class="sxs-lookup"><span data-stu-id="cf0bd-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf0bd-470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cf0bd-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf0bd-471">1.0</span><span class="sxs-lookup"><span data-stu-id="cf0bd-471">1.0</span></span>|
|[<span data-ttu-id="cf0bd-472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cf0bd-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf0bd-473">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="cf0bd-473">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="cf0bd-474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cf0bd-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf0bd-475">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cf0bd-475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf0bd-476">示例</span><span class="sxs-lookup"><span data-stu-id="cf0bd-476">Example</span></span>

<span data-ttu-id="cf0bd-477">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="cf0bd-477">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
