---
title: "\"Context.subname\"-\"邮箱-要求集 1.3\""
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 05b7d82e036cc29526c18bf97c6a1472778c1959
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696230"
---
# <a name="mailbox"></a><span data-ttu-id="95173-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="95173-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="95173-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="95173-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="95173-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="95173-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="95173-105">要求</span><span class="sxs-lookup"><span data-stu-id="95173-105">Requirements</span></span>

|<span data-ttu-id="95173-106">要求</span><span class="sxs-lookup"><span data-stu-id="95173-106">Requirement</span></span>| <span data-ttu-id="95173-107">值</span><span class="sxs-lookup"><span data-stu-id="95173-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-109">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-109">1.0</span></span>|
|[<span data-ttu-id="95173-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-111">受限</span><span class="sxs-lookup"><span data-stu-id="95173-111">Restricted</span></span>|
|[<span data-ttu-id="95173-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-113">Compose or Read</span></span>|

<span data-ttu-id="95173-114">| [mailbox.ewsurl](#ewsurl-string) |Member | |[office.context.mailbox.converttoewsid](#converttoewsiditemid-restversion--string) |方法 | |[convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) |方法 | |[office.context.mailbox.converttorestid](#converttorestiditemid-restversion--string) |方法 | |[convertToUtcClientTime](#converttoutcclienttimeinput--date) |方法 | |[displayAppointmentForm](#displayappointmentformitemid) |方法 | |[displayMessageForm](#displaymessageformitemid) |方法 | |[displayNewAppointmentForm](#displaynewappointmentformparameters) |方法 | |[mailbox.getcallbacktokenasync](#getcallbacktokenasynccallback-usercontext) |方法 | |[getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) |方法 | |[makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) |方法 |</span><span class="sxs-lookup"><span data-stu-id="95173-114">| [ewsUrl](#ewsurl-string) | Member | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Method | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Method | | [convertToRestId](#converttorestiditemid-restversion--string) | Method | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Method | | [displayAppointmentForm](#displayappointmentformitemid) | Method | | [displayMessageForm](#displaymessageformitemid) | Method | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Method | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Method | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Method | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Method |</span></span>

### <a name="namespaces"></a><span data-ttu-id="95173-115">命名空间</span><span class="sxs-lookup"><span data-stu-id="95173-115">Namespaces</span></span>

<span data-ttu-id="95173-116">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="95173-116">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="95173-117">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="95173-117">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="95173-118">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="95173-118">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="95173-119">成员</span><span class="sxs-lookup"><span data-stu-id="95173-119">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="95173-120">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="95173-120">ewsUrl: String</span></span>

<span data-ttu-id="95173-121">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="95173-121">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="95173-122">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="95173-122">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-123">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="95173-123">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="95173-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="95173-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="95173-126">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="95173-126">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="95173-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="95173-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="95173-129">类型</span><span class="sxs-lookup"><span data-stu-id="95173-129">Type</span></span>

*   <span data-ttu-id="95173-130">String</span><span class="sxs-lookup"><span data-stu-id="95173-130">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95173-131">要求</span><span class="sxs-lookup"><span data-stu-id="95173-131">Requirements</span></span>

|<span data-ttu-id="95173-132">要求</span><span class="sxs-lookup"><span data-stu-id="95173-132">Requirement</span></span>| <span data-ttu-id="95173-133">值</span><span class="sxs-lookup"><span data-stu-id="95173-133">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-134">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-135">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-135">1.0</span></span>|
|[<span data-ttu-id="95173-136">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-136">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-137">ReadItem</span></span>|
|[<span data-ttu-id="95173-138">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-138">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-139">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-139">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="95173-140">方法</span><span class="sxs-lookup"><span data-stu-id="95173-140">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="95173-141">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="95173-141">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="95173-142">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="95173-142">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-143">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="95173-143">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="95173-p104">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="95173-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-146">参数</span><span class="sxs-lookup"><span data-stu-id="95173-146">Parameters</span></span>

|<span data-ttu-id="95173-147">名称</span><span class="sxs-lookup"><span data-stu-id="95173-147">Name</span></span>| <span data-ttu-id="95173-148">类型</span><span class="sxs-lookup"><span data-stu-id="95173-148">Type</span></span>| <span data-ttu-id="95173-149">说明</span><span class="sxs-lookup"><span data-stu-id="95173-149">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="95173-150">String</span><span class="sxs-lookup"><span data-stu-id="95173-150">String</span></span>|<span data-ttu-id="95173-151">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="95173-151">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="95173-152">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="95173-152">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="95173-153">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="95173-153">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-154">要求</span><span class="sxs-lookup"><span data-stu-id="95173-154">Requirements</span></span>

|<span data-ttu-id="95173-155">要求</span><span class="sxs-lookup"><span data-stu-id="95173-155">Requirement</span></span>| <span data-ttu-id="95173-156">值</span><span class="sxs-lookup"><span data-stu-id="95173-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-157">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-158">1.3</span><span class="sxs-lookup"><span data-stu-id="95173-158">1.3</span></span>|
|[<span data-ttu-id="95173-159">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-159">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-160">受限</span><span class="sxs-lookup"><span data-stu-id="95173-160">Restricted</span></span>|
|[<span data-ttu-id="95173-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-162">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95173-163">返回：</span><span class="sxs-lookup"><span data-stu-id="95173-163">Returns:</span></span>

<span data-ttu-id="95173-164">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="95173-164">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="95173-165">示例</span><span class="sxs-lookup"><span data-stu-id="95173-165">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="95173-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="95173-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="95173-167">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="95173-167">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="95173-168">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="95173-168">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="95173-169">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="95173-169">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="95173-170">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="95173-170">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="95173-171">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="95173-171">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="95173-172">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="95173-172">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-173">参数</span><span class="sxs-lookup"><span data-stu-id="95173-173">Parameters</span></span>

|<span data-ttu-id="95173-174">名称</span><span class="sxs-lookup"><span data-stu-id="95173-174">Name</span></span>| <span data-ttu-id="95173-175">类型</span><span class="sxs-lookup"><span data-stu-id="95173-175">Type</span></span>| <span data-ttu-id="95173-176">描述</span><span class="sxs-lookup"><span data-stu-id="95173-176">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="95173-177">日期</span><span class="sxs-lookup"><span data-stu-id="95173-177">Date</span></span>|<span data-ttu-id="95173-178">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="95173-178">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-179">要求</span><span class="sxs-lookup"><span data-stu-id="95173-179">Requirements</span></span>

|<span data-ttu-id="95173-180">要求</span><span class="sxs-lookup"><span data-stu-id="95173-180">Requirement</span></span>| <span data-ttu-id="95173-181">值</span><span class="sxs-lookup"><span data-stu-id="95173-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-183">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-183">1.0</span></span>|
|[<span data-ttu-id="95173-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-185">ReadItem</span></span>|
|[<span data-ttu-id="95173-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-187">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95173-188">返回：</span><span class="sxs-lookup"><span data-stu-id="95173-188">Returns:</span></span>

<span data-ttu-id="95173-189">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="95173-189">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="95173-190">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="95173-190">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="95173-191">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="95173-191">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-192">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="95173-192">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="95173-p107">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="95173-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-195">参数</span><span class="sxs-lookup"><span data-stu-id="95173-195">Parameters</span></span>

|<span data-ttu-id="95173-196">名称</span><span class="sxs-lookup"><span data-stu-id="95173-196">Name</span></span>| <span data-ttu-id="95173-197">类型</span><span class="sxs-lookup"><span data-stu-id="95173-197">Type</span></span>| <span data-ttu-id="95173-198">说明</span><span class="sxs-lookup"><span data-stu-id="95173-198">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="95173-199">字符串</span><span class="sxs-lookup"><span data-stu-id="95173-199">String</span></span>|<span data-ttu-id="95173-200">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="95173-200">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="95173-201">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="95173-201">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="95173-202">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="95173-202">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-203">要求</span><span class="sxs-lookup"><span data-stu-id="95173-203">Requirements</span></span>

|<span data-ttu-id="95173-204">要求</span><span class="sxs-lookup"><span data-stu-id="95173-204">Requirement</span></span>| <span data-ttu-id="95173-205">值</span><span class="sxs-lookup"><span data-stu-id="95173-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-207">1.3</span><span class="sxs-lookup"><span data-stu-id="95173-207">1.3</span></span>|
|[<span data-ttu-id="95173-208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-209">受限</span><span class="sxs-lookup"><span data-stu-id="95173-209">Restricted</span></span>|
|[<span data-ttu-id="95173-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95173-212">返回：</span><span class="sxs-lookup"><span data-stu-id="95173-212">Returns:</span></span>

<span data-ttu-id="95173-213">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="95173-213">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="95173-214">示例</span><span class="sxs-lookup"><span data-stu-id="95173-214">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="95173-215">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="95173-215">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="95173-216">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-216">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="95173-217">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-217">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-218">参数</span><span class="sxs-lookup"><span data-stu-id="95173-218">Parameters</span></span>

|<span data-ttu-id="95173-219">名称</span><span class="sxs-lookup"><span data-stu-id="95173-219">Name</span></span>| <span data-ttu-id="95173-220">类型</span><span class="sxs-lookup"><span data-stu-id="95173-220">Type</span></span>| <span data-ttu-id="95173-221">说明</span><span class="sxs-lookup"><span data-stu-id="95173-221">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="95173-222">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="95173-222">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="95173-223">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="95173-223">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-224">要求</span><span class="sxs-lookup"><span data-stu-id="95173-224">Requirements</span></span>

|<span data-ttu-id="95173-225">要求</span><span class="sxs-lookup"><span data-stu-id="95173-225">Requirement</span></span>| <span data-ttu-id="95173-226">值</span><span class="sxs-lookup"><span data-stu-id="95173-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-228">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-228">1.0</span></span>|
|[<span data-ttu-id="95173-229">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-230">ReadItem</span></span>|
|[<span data-ttu-id="95173-231">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-232">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-232">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="95173-233">返回：</span><span class="sxs-lookup"><span data-stu-id="95173-233">Returns:</span></span>

<span data-ttu-id="95173-234">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-234">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="95173-235">类型: Date</span><span class="sxs-lookup"><span data-stu-id="95173-235">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="95173-236">示例</span><span class="sxs-lookup"><span data-stu-id="95173-236">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="95173-237">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="95173-237">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="95173-238">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="95173-238">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-239">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="95173-239">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="95173-240">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="95173-240">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="95173-241">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="95173-241">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="95173-242">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="95173-242">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="95173-243">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="95173-243">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="95173-244">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="95173-244">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-245">参数</span><span class="sxs-lookup"><span data-stu-id="95173-245">Parameters</span></span>

|<span data-ttu-id="95173-246">名称</span><span class="sxs-lookup"><span data-stu-id="95173-246">Name</span></span>| <span data-ttu-id="95173-247">类型</span><span class="sxs-lookup"><span data-stu-id="95173-247">Type</span></span>| <span data-ttu-id="95173-248">说明</span><span class="sxs-lookup"><span data-stu-id="95173-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="95173-249">String</span><span class="sxs-lookup"><span data-stu-id="95173-249">String</span></span>|<span data-ttu-id="95173-250">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="95173-250">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-251">要求</span><span class="sxs-lookup"><span data-stu-id="95173-251">Requirements</span></span>

|<span data-ttu-id="95173-252">要求</span><span class="sxs-lookup"><span data-stu-id="95173-252">Requirement</span></span>| <span data-ttu-id="95173-253">值</span><span class="sxs-lookup"><span data-stu-id="95173-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-254">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-255">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-255">1.0</span></span>|
|[<span data-ttu-id="95173-256">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-257">ReadItem</span></span>|
|[<span data-ttu-id="95173-258">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-259">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95173-260">示例</span><span class="sxs-lookup"><span data-stu-id="95173-260">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="95173-261">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="95173-261">displayMessageForm(itemId)</span></span>

<span data-ttu-id="95173-262">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="95173-262">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-263">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="95173-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="95173-264">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="95173-264">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="95173-265">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="95173-265">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="95173-266">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="95173-266">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="95173-p109">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="95173-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-269">参数</span><span class="sxs-lookup"><span data-stu-id="95173-269">Parameters</span></span>

|<span data-ttu-id="95173-270">名称</span><span class="sxs-lookup"><span data-stu-id="95173-270">Name</span></span>| <span data-ttu-id="95173-271">类型</span><span class="sxs-lookup"><span data-stu-id="95173-271">Type</span></span>| <span data-ttu-id="95173-272">说明</span><span class="sxs-lookup"><span data-stu-id="95173-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="95173-273">String</span><span class="sxs-lookup"><span data-stu-id="95173-273">String</span></span>|<span data-ttu-id="95173-274">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="95173-274">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-275">要求</span><span class="sxs-lookup"><span data-stu-id="95173-275">Requirements</span></span>

|<span data-ttu-id="95173-276">要求</span><span class="sxs-lookup"><span data-stu-id="95173-276">Requirement</span></span>| <span data-ttu-id="95173-277">值</span><span class="sxs-lookup"><span data-stu-id="95173-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-279">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-279">1.0</span></span>|
|[<span data-ttu-id="95173-280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-281">ReadItem</span></span>|
|[<span data-ttu-id="95173-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-283">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95173-284">示例</span><span class="sxs-lookup"><span data-stu-id="95173-284">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="95173-285">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="95173-285">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="95173-286">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="95173-286">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-287">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="95173-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="95173-p110">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="95173-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="95173-290">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="95173-290">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="95173-291">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="95173-291">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="95173-292">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="95173-292">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="95173-p112">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="95173-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="95173-295">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="95173-295">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-296">参数</span><span class="sxs-lookup"><span data-stu-id="95173-296">Parameters</span></span>

|<span data-ttu-id="95173-297">名称</span><span class="sxs-lookup"><span data-stu-id="95173-297">Name</span></span>| <span data-ttu-id="95173-298">类型</span><span class="sxs-lookup"><span data-stu-id="95173-298">Type</span></span>| <span data-ttu-id="95173-299">说明</span><span class="sxs-lookup"><span data-stu-id="95173-299">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="95173-300">对象</span><span class="sxs-lookup"><span data-stu-id="95173-300">Object</span></span> | <span data-ttu-id="95173-301">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="95173-301">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="95173-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="95173-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="95173-p113">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="95173-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="95173-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="95173-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="95173-p114">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="95173-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="95173-308">日期</span><span class="sxs-lookup"><span data-stu-id="95173-308">Date</span></span> | <span data-ttu-id="95173-309">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-309">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="95173-310">Date</span><span class="sxs-lookup"><span data-stu-id="95173-310">Date</span></span> | <span data-ttu-id="95173-311">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-311">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="95173-312">String</span><span class="sxs-lookup"><span data-stu-id="95173-312">String</span></span> | <span data-ttu-id="95173-p115">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="95173-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="95173-315">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="95173-315">Array.&lt;String&gt;</span></span> | <span data-ttu-id="95173-p116">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="95173-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="95173-318">String</span><span class="sxs-lookup"><span data-stu-id="95173-318">String</span></span> | <span data-ttu-id="95173-p117">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="95173-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="95173-321">字符串</span><span class="sxs-lookup"><span data-stu-id="95173-321">String</span></span> | <span data-ttu-id="95173-p118">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="95173-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="95173-324">要求</span><span class="sxs-lookup"><span data-stu-id="95173-324">Requirements</span></span>

|<span data-ttu-id="95173-325">要求</span><span class="sxs-lookup"><span data-stu-id="95173-325">Requirement</span></span>| <span data-ttu-id="95173-326">值</span><span class="sxs-lookup"><span data-stu-id="95173-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-328">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-328">1.0</span></span>|
|[<span data-ttu-id="95173-329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-330">ReadItem</span></span>|
|[<span data-ttu-id="95173-331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-332">阅读</span><span class="sxs-lookup"><span data-stu-id="95173-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95173-333">示例</span><span class="sxs-lookup"><span data-stu-id="95173-333">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="95173-334">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="95173-334">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="95173-335">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="95173-335">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="95173-p119">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="95173-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="95173-p120">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="95173-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="95173-341">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="95173-341">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="95173-p121">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="95173-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-344">参数</span><span class="sxs-lookup"><span data-stu-id="95173-344">Parameters</span></span>

|<span data-ttu-id="95173-345">名称</span><span class="sxs-lookup"><span data-stu-id="95173-345">Name</span></span>| <span data-ttu-id="95173-346">类型</span><span class="sxs-lookup"><span data-stu-id="95173-346">Type</span></span>| <span data-ttu-id="95173-347">属性</span><span class="sxs-lookup"><span data-stu-id="95173-347">Attributes</span></span>| <span data-ttu-id="95173-348">说明</span><span class="sxs-lookup"><span data-stu-id="95173-348">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="95173-349">function</span><span class="sxs-lookup"><span data-stu-id="95173-349">function</span></span>||<span data-ttu-id="95173-350">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="95173-350">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95173-351">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="95173-351">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="95173-352">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="95173-352">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="95173-353">对象</span><span class="sxs-lookup"><span data-stu-id="95173-353">Object</span></span>| <span data-ttu-id="95173-354">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="95173-354">&lt;optional&gt;</span></span>|<span data-ttu-id="95173-355">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="95173-355">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="95173-356">错误</span><span class="sxs-lookup"><span data-stu-id="95173-356">Errors</span></span>

|<span data-ttu-id="95173-357">错误代码</span><span class="sxs-lookup"><span data-stu-id="95173-357">Error code</span></span>|<span data-ttu-id="95173-358">说明</span><span class="sxs-lookup"><span data-stu-id="95173-358">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="95173-359">请求失败。</span><span class="sxs-lookup"><span data-stu-id="95173-359">The request has failed.</span></span> <span data-ttu-id="95173-360">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-360">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="95173-361">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="95173-361">The Exchange server returned an error.</span></span> <span data-ttu-id="95173-362">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-362">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="95173-363">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="95173-363">The user is no longer connected to the network.</span></span> <span data-ttu-id="95173-364">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="95173-364">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-365">要求</span><span class="sxs-lookup"><span data-stu-id="95173-365">Requirements</span></span>

|<span data-ttu-id="95173-366">要求</span><span class="sxs-lookup"><span data-stu-id="95173-366">Requirement</span></span>| <span data-ttu-id="95173-367">值</span><span class="sxs-lookup"><span data-stu-id="95173-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-368">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-369">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-369">1.0</span></span>|
|[<span data-ttu-id="95173-370">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-371">ReadItem</span></span>|
|[<span data-ttu-id="95173-372">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-373">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="95173-373">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="95173-374">示例</span><span class="sxs-lookup"><span data-stu-id="95173-374">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="95173-375">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="95173-375">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="95173-376">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="95173-376">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="95173-377">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="95173-377">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-378">参数</span><span class="sxs-lookup"><span data-stu-id="95173-378">Parameters</span></span>

|<span data-ttu-id="95173-379">名称</span><span class="sxs-lookup"><span data-stu-id="95173-379">Name</span></span>| <span data-ttu-id="95173-380">类型</span><span class="sxs-lookup"><span data-stu-id="95173-380">Type</span></span>| <span data-ttu-id="95173-381">属性</span><span class="sxs-lookup"><span data-stu-id="95173-381">Attributes</span></span>| <span data-ttu-id="95173-382">说明</span><span class="sxs-lookup"><span data-stu-id="95173-382">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="95173-383">function</span><span class="sxs-lookup"><span data-stu-id="95173-383">function</span></span>||<span data-ttu-id="95173-384">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="95173-384">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95173-385">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="95173-385">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="95173-386">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="95173-386">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="95173-387">对象</span><span class="sxs-lookup"><span data-stu-id="95173-387">Object</span></span>| <span data-ttu-id="95173-388">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="95173-388">&lt;optional&gt;</span></span>|<span data-ttu-id="95173-389">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="95173-389">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="95173-390">错误</span><span class="sxs-lookup"><span data-stu-id="95173-390">Errors</span></span>

|<span data-ttu-id="95173-391">错误代码</span><span class="sxs-lookup"><span data-stu-id="95173-391">Error code</span></span>|<span data-ttu-id="95173-392">说明</span><span class="sxs-lookup"><span data-stu-id="95173-392">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="95173-393">请求失败。</span><span class="sxs-lookup"><span data-stu-id="95173-393">The request has failed.</span></span> <span data-ttu-id="95173-394">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-394">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="95173-395">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="95173-395">The Exchange server returned an error.</span></span> <span data-ttu-id="95173-396">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="95173-396">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="95173-397">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="95173-397">The user is no longer connected to the network.</span></span> <span data-ttu-id="95173-398">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="95173-398">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-399">要求</span><span class="sxs-lookup"><span data-stu-id="95173-399">Requirements</span></span>

|<span data-ttu-id="95173-400">要求</span><span class="sxs-lookup"><span data-stu-id="95173-400">Requirement</span></span>| <span data-ttu-id="95173-401">值</span><span class="sxs-lookup"><span data-stu-id="95173-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-402">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-403">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-403">1.0</span></span>|
|[<span data-ttu-id="95173-404">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="95173-405">ReadItem</span></span>|
|[<span data-ttu-id="95173-406">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-407">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-407">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95173-408">示例</span><span class="sxs-lookup"><span data-stu-id="95173-408">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="95173-409">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="95173-409">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="95173-410">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="95173-410">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="95173-411">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="95173-411">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="95173-412">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="95173-412">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="95173-413">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="95173-413">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="95173-414">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="95173-414">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="95173-415">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="95173-415">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="95173-416">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="95173-416">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="95173-417">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="95173-417">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="95173-418">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="95173-418">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="95173-p129">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="95173-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="95173-421">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="95173-421">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="95173-422">版本差异</span><span class="sxs-lookup"><span data-stu-id="95173-422">Version differences</span></span>

<span data-ttu-id="95173-423">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="95173-423">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="95173-p130">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="95173-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="95173-427">参数</span><span class="sxs-lookup"><span data-stu-id="95173-427">Parameters</span></span>

|<span data-ttu-id="95173-428">名称</span><span class="sxs-lookup"><span data-stu-id="95173-428">Name</span></span>| <span data-ttu-id="95173-429">类型</span><span class="sxs-lookup"><span data-stu-id="95173-429">Type</span></span>| <span data-ttu-id="95173-430">属性</span><span class="sxs-lookup"><span data-stu-id="95173-430">Attributes</span></span>| <span data-ttu-id="95173-431">说明</span><span class="sxs-lookup"><span data-stu-id="95173-431">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="95173-432">字符串</span><span class="sxs-lookup"><span data-stu-id="95173-432">String</span></span>||<span data-ttu-id="95173-433">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="95173-433">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="95173-434">函数</span><span class="sxs-lookup"><span data-stu-id="95173-434">function</span></span>||<span data-ttu-id="95173-435">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="95173-435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="95173-436">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="95173-436">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="95173-437">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="95173-437">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="95173-438">对象</span><span class="sxs-lookup"><span data-stu-id="95173-438">Object</span></span>| <span data-ttu-id="95173-439">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="95173-439">&lt;optional&gt;</span></span>|<span data-ttu-id="95173-440">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="95173-440">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95173-441">要求</span><span class="sxs-lookup"><span data-stu-id="95173-441">Requirements</span></span>

|<span data-ttu-id="95173-442">要求</span><span class="sxs-lookup"><span data-stu-id="95173-442">Requirement</span></span>| <span data-ttu-id="95173-443">值</span><span class="sxs-lookup"><span data-stu-id="95173-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="95173-444">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95173-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="95173-445">1.0</span><span class="sxs-lookup"><span data-stu-id="95173-445">1.0</span></span>|
|[<span data-ttu-id="95173-446">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="95173-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="95173-447">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="95173-447">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="95173-448">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95173-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="95173-449">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95173-449">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95173-450">示例</span><span class="sxs-lookup"><span data-stu-id="95173-450">Example</span></span>

<span data-ttu-id="95173-451">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="95173-451">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
