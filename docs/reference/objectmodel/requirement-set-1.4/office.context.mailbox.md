---
title: "\"Context.subname\"-\"邮箱-要求集 1.4\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 373ac1f4361ae94f788a61e49c304a1cfdb90232
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064702"
---
# <a name="mailbox"></a><span data-ttu-id="167a8-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="167a8-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="167a8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="167a8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="167a8-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="167a8-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167a8-105">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-105">Requirements</span></span>

|<span data-ttu-id="167a8-106">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-106">Requirement</span></span>| <span data-ttu-id="167a8-107">值</span><span class="sxs-lookup"><span data-stu-id="167a8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-109">1.0</span></span>|
|[<span data-ttu-id="167a8-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-111">受限</span><span class="sxs-lookup"><span data-stu-id="167a8-111">Restricted</span></span>|
|[<span data-ttu-id="167a8-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="167a8-114">命名空间</span><span class="sxs-lookup"><span data-stu-id="167a8-114">Namespaces</span></span>

<span data-ttu-id="167a8-115">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="167a8-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="167a8-116">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="167a8-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="167a8-117">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="167a8-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="167a8-118">成员</span><span class="sxs-lookup"><span data-stu-id="167a8-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="167a8-119">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="167a8-119">ewsUrl: String</span></span>

<span data-ttu-id="167a8-120">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="167a8-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="167a8-121">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="167a8-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-122">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="167a8-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="167a8-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="167a8-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="167a8-125">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="167a8-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="167a8-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="167a8-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="167a8-128">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-128">Type</span></span>

*   <span data-ttu-id="167a8-129">String</span><span class="sxs-lookup"><span data-stu-id="167a8-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167a8-130">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-130">Requirements</span></span>

|<span data-ttu-id="167a8-131">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-131">Requirement</span></span>| <span data-ttu-id="167a8-132">值</span><span class="sxs-lookup"><span data-stu-id="167a8-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-133">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-134">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-134">1.0</span></span>|
|[<span data-ttu-id="167a8-135">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-136">ReadItem</span></span>|
|[<span data-ttu-id="167a8-137">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-138">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="167a8-139">方法</span><span class="sxs-lookup"><span data-stu-id="167a8-139">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="167a8-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="167a8-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="167a8-141">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="167a8-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-142">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="167a8-142">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="167a8-p104">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="167a8-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-145">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-145">Parameters</span></span>

|<span data-ttu-id="167a8-146">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-146">Name</span></span>| <span data-ttu-id="167a8-147">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-147">Type</span></span>| <span data-ttu-id="167a8-148">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="167a8-149">String</span><span class="sxs-lookup"><span data-stu-id="167a8-149">String</span></span>|<span data-ttu-id="167a8-150">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="167a8-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="167a8-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="167a8-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="167a8-152">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="167a8-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-153">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-153">Requirements</span></span>

|<span data-ttu-id="167a8-154">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-154">Requirement</span></span>| <span data-ttu-id="167a8-155">值</span><span class="sxs-lookup"><span data-stu-id="167a8-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-156">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-157">1.3</span><span class="sxs-lookup"><span data-stu-id="167a8-157">1.3</span></span>|
|[<span data-ttu-id="167a8-158">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-158">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-159">受限</span><span class="sxs-lookup"><span data-stu-id="167a8-159">Restricted</span></span>|
|[<span data-ttu-id="167a8-160">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-160">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-161">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167a8-162">返回：</span><span class="sxs-lookup"><span data-stu-id="167a8-162">Returns:</span></span>

<span data-ttu-id="167a8-163">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="167a8-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="167a8-164">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="167a8-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="167a8-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="167a8-166">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="167a8-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="167a8-167">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="167a8-167">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="167a8-168">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="167a8-168">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="167a8-169">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="167a8-169">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="167a8-170">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="167a8-170">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="167a8-171">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="167a8-171">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-172">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-172">Parameters</span></span>

|<span data-ttu-id="167a8-173">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-173">Name</span></span>| <span data-ttu-id="167a8-174">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-174">Type</span></span>| <span data-ttu-id="167a8-175">描述</span><span class="sxs-lookup"><span data-stu-id="167a8-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="167a8-176">日期</span><span class="sxs-lookup"><span data-stu-id="167a8-176">Date</span></span>|<span data-ttu-id="167a8-177">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="167a8-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-178">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-178">Requirements</span></span>

|<span data-ttu-id="167a8-179">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-179">Requirement</span></span>| <span data-ttu-id="167a8-180">值</span><span class="sxs-lookup"><span data-stu-id="167a8-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-182">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-182">1.0</span></span>|
|[<span data-ttu-id="167a8-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-184">ReadItem</span></span>|
|[<span data-ttu-id="167a8-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167a8-187">返回：</span><span class="sxs-lookup"><span data-stu-id="167a8-187">Returns:</span></span>

<span data-ttu-id="167a8-188">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="167a8-188">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="167a8-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="167a8-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="167a8-190">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="167a8-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-191">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="167a8-191">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="167a8-p107">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="167a8-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-194">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-194">Parameters</span></span>

|<span data-ttu-id="167a8-195">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-195">Name</span></span>| <span data-ttu-id="167a8-196">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-196">Type</span></span>| <span data-ttu-id="167a8-197">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="167a8-198">字符串</span><span class="sxs-lookup"><span data-stu-id="167a8-198">String</span></span>|<span data-ttu-id="167a8-199">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="167a8-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="167a8-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="167a8-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="167a8-201">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="167a8-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-202">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-202">Requirements</span></span>

|<span data-ttu-id="167a8-203">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-203">Requirement</span></span>| <span data-ttu-id="167a8-204">值</span><span class="sxs-lookup"><span data-stu-id="167a8-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-206">1.3</span><span class="sxs-lookup"><span data-stu-id="167a8-206">1.3</span></span>|
|[<span data-ttu-id="167a8-207">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-207">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-208">受限</span><span class="sxs-lookup"><span data-stu-id="167a8-208">Restricted</span></span>|
|[<span data-ttu-id="167a8-209">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-209">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-210">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167a8-211">返回：</span><span class="sxs-lookup"><span data-stu-id="167a8-211">Returns:</span></span>

<span data-ttu-id="167a8-212">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="167a8-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="167a8-213">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="167a8-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="167a8-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="167a8-215">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="167a8-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="167a8-216">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="167a8-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-217">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-217">Parameters</span></span>

|<span data-ttu-id="167a8-218">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-218">Name</span></span>| <span data-ttu-id="167a8-219">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-219">Type</span></span>| <span data-ttu-id="167a8-220">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="167a8-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="167a8-221">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="167a8-222">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="167a8-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-223">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-223">Requirements</span></span>

|<span data-ttu-id="167a8-224">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-224">Requirement</span></span>| <span data-ttu-id="167a8-225">值</span><span class="sxs-lookup"><span data-stu-id="167a8-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-226">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-227">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-227">1.0</span></span>|
|[<span data-ttu-id="167a8-228">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-228">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-229">ReadItem</span></span>|
|[<span data-ttu-id="167a8-230">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-230">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-231">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167a8-232">返回：</span><span class="sxs-lookup"><span data-stu-id="167a8-232">Returns:</span></span>

<span data-ttu-id="167a8-233">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="167a8-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="167a8-234">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="167a8-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="167a8-235">日期</span><span class="sxs-lookup"><span data-stu-id="167a8-235">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="167a8-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="167a8-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="167a8-237">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="167a8-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-238">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="167a8-238">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="167a8-239">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="167a8-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="167a8-240">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="167a8-240">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="167a8-241">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="167a8-241">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="167a8-242">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="167a8-242">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="167a8-243">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="167a8-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-244">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-244">Parameters</span></span>

|<span data-ttu-id="167a8-245">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-245">Name</span></span>| <span data-ttu-id="167a8-246">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-246">Type</span></span>| <span data-ttu-id="167a8-247">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="167a8-248">String</span><span class="sxs-lookup"><span data-stu-id="167a8-248">String</span></span>|<span data-ttu-id="167a8-249">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="167a8-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-250">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-250">Requirements</span></span>

|<span data-ttu-id="167a8-251">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-251">Requirement</span></span>| <span data-ttu-id="167a8-252">值</span><span class="sxs-lookup"><span data-stu-id="167a8-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-254">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-254">1.0</span></span>|
|[<span data-ttu-id="167a8-255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-256">ReadItem</span></span>|
|[<span data-ttu-id="167a8-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-258">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167a8-259">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="167a8-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="167a8-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="167a8-261">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="167a8-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-262">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="167a8-262">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="167a8-263">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="167a8-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="167a8-264">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="167a8-264">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="167a8-265">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="167a8-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="167a8-p109">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="167a8-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-268">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-268">Parameters</span></span>

|<span data-ttu-id="167a8-269">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-269">Name</span></span>| <span data-ttu-id="167a8-270">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-270">Type</span></span>| <span data-ttu-id="167a8-271">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="167a8-272">String</span><span class="sxs-lookup"><span data-stu-id="167a8-272">String</span></span>|<span data-ttu-id="167a8-273">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="167a8-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-274">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-274">Requirements</span></span>

|<span data-ttu-id="167a8-275">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-275">Requirement</span></span>| <span data-ttu-id="167a8-276">值</span><span class="sxs-lookup"><span data-stu-id="167a8-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-277">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-278">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-278">1.0</span></span>|
|[<span data-ttu-id="167a8-279">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-280">ReadItem</span></span>|
|[<span data-ttu-id="167a8-281">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-282">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167a8-283">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="167a8-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="167a8-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="167a8-285">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="167a8-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-286">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="167a8-286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="167a8-p110">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="167a8-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="167a8-289">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="167a8-289">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="167a8-290">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="167a8-290">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="167a8-291">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="167a8-291">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="167a8-p112">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="167a8-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="167a8-294">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="167a8-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-295">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-295">Parameters</span></span>

|<span data-ttu-id="167a8-296">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-296">Name</span></span>| <span data-ttu-id="167a8-297">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-297">Type</span></span>| <span data-ttu-id="167a8-298">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="167a8-299">对象</span><span class="sxs-lookup"><span data-stu-id="167a8-299">Object</span></span> | <span data-ttu-id="167a8-300">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="167a8-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="167a8-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="167a8-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="167a8-p113">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="167a8-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="167a8-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="167a8-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="167a8-p114">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="167a8-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="167a8-307">日期</span><span class="sxs-lookup"><span data-stu-id="167a8-307">Date</span></span> | <span data-ttu-id="167a8-308">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="167a8-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="167a8-309">Date</span><span class="sxs-lookup"><span data-stu-id="167a8-309">Date</span></span> | <span data-ttu-id="167a8-310">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="167a8-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="167a8-311">String</span><span class="sxs-lookup"><span data-stu-id="167a8-311">String</span></span> | <span data-ttu-id="167a8-p115">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="167a8-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="167a8-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="167a8-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="167a8-p116">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="167a8-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="167a8-317">String</span><span class="sxs-lookup"><span data-stu-id="167a8-317">String</span></span> | <span data-ttu-id="167a8-p117">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="167a8-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="167a8-320">字符串</span><span class="sxs-lookup"><span data-stu-id="167a8-320">String</span></span> | <span data-ttu-id="167a8-p118">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="167a8-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="167a8-323">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-323">Requirements</span></span>

|<span data-ttu-id="167a8-324">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-324">Requirement</span></span>| <span data-ttu-id="167a8-325">值</span><span class="sxs-lookup"><span data-stu-id="167a8-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-326">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-327">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-327">1.0</span></span>|
|[<span data-ttu-id="167a8-328">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-329">ReadItem</span></span>|
|[<span data-ttu-id="167a8-330">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-331">阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167a8-332">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="167a8-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="167a8-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="167a8-334">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="167a8-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="167a8-p119">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="167a8-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="167a8-p120">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="167a8-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="167a8-340">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="167a8-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="167a8-p121">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="167a8-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-343">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-343">Parameters</span></span>

|<span data-ttu-id="167a8-344">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-344">Name</span></span>| <span data-ttu-id="167a8-345">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-345">Type</span></span>| <span data-ttu-id="167a8-346">属性</span><span class="sxs-lookup"><span data-stu-id="167a8-346">Attributes</span></span>| <span data-ttu-id="167a8-347">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="167a8-348">函数</span><span class="sxs-lookup"><span data-stu-id="167a8-348">function</span></span>||<span data-ttu-id="167a8-p122">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="167a8-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="167a8-351">对象</span><span class="sxs-lookup"><span data-stu-id="167a8-351">Object</span></span>| <span data-ttu-id="167a8-352">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="167a8-352">&lt;optional&gt;</span></span>|<span data-ttu-id="167a8-353">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="167a8-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-354">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-354">Requirements</span></span>

|<span data-ttu-id="167a8-355">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-355">Requirement</span></span>| <span data-ttu-id="167a8-356">值</span><span class="sxs-lookup"><span data-stu-id="167a8-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-357">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-358">1.3</span><span class="sxs-lookup"><span data-stu-id="167a8-358">1.3</span></span>|
|[<span data-ttu-id="167a8-359">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-360">ReadItem</span></span>|
|[<span data-ttu-id="167a8-361">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-362">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="167a8-363">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="167a8-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="167a8-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="167a8-365">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="167a8-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="167a8-366">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="167a8-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-367">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-367">Parameters</span></span>

|<span data-ttu-id="167a8-368">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-368">Name</span></span>| <span data-ttu-id="167a8-369">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-369">Type</span></span>| <span data-ttu-id="167a8-370">属性</span><span class="sxs-lookup"><span data-stu-id="167a8-370">Attributes</span></span>| <span data-ttu-id="167a8-371">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="167a8-372">function</span><span class="sxs-lookup"><span data-stu-id="167a8-372">function</span></span>||<span data-ttu-id="167a8-373">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="167a8-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="167a8-374">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="167a8-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="167a8-375">对象</span><span class="sxs-lookup"><span data-stu-id="167a8-375">Object</span></span>| <span data-ttu-id="167a8-376">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="167a8-376">&lt;optional&gt;</span></span>|<span data-ttu-id="167a8-377">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="167a8-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-378">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-378">Requirements</span></span>

|<span data-ttu-id="167a8-379">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-379">Requirement</span></span>| <span data-ttu-id="167a8-380">值</span><span class="sxs-lookup"><span data-stu-id="167a8-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-381">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-382">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-382">1.0</span></span>|
|[<span data-ttu-id="167a8-383">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167a8-384">ReadItem</span></span>|
|[<span data-ttu-id="167a8-385">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-386">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167a8-387">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="167a8-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="167a8-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="167a8-389">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="167a8-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-390">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="167a8-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="167a8-391">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="167a8-391">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="167a8-392">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="167a8-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="167a8-393">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="167a8-393">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="167a8-394">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="167a8-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="167a8-395">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="167a8-395">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="167a8-396">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="167a8-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="167a8-397">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="167a8-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="167a8-p124">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="167a8-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="167a8-400">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="167a8-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="167a8-401">版本差异</span><span class="sxs-lookup"><span data-stu-id="167a8-401">Version differences</span></span>

<span data-ttu-id="167a8-402">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="167a8-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="167a8-p125">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="167a8-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167a8-406">参数</span><span class="sxs-lookup"><span data-stu-id="167a8-406">Parameters</span></span>

|<span data-ttu-id="167a8-407">名称</span><span class="sxs-lookup"><span data-stu-id="167a8-407">Name</span></span>| <span data-ttu-id="167a8-408">类型</span><span class="sxs-lookup"><span data-stu-id="167a8-408">Type</span></span>| <span data-ttu-id="167a8-409">属性</span><span class="sxs-lookup"><span data-stu-id="167a8-409">Attributes</span></span>| <span data-ttu-id="167a8-410">说明</span><span class="sxs-lookup"><span data-stu-id="167a8-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="167a8-411">字符串</span><span class="sxs-lookup"><span data-stu-id="167a8-411">String</span></span>||<span data-ttu-id="167a8-412">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="167a8-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="167a8-413">function</span><span class="sxs-lookup"><span data-stu-id="167a8-413">function</span></span>||<span data-ttu-id="167a8-414">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="167a8-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="167a8-415">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="167a8-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="167a8-416">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="167a8-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="167a8-417">对象</span><span class="sxs-lookup"><span data-stu-id="167a8-417">Object</span></span>| <span data-ttu-id="167a8-418">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="167a8-418">&lt;optional&gt;</span></span>|<span data-ttu-id="167a8-419">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="167a8-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167a8-420">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-420">Requirements</span></span>

|<span data-ttu-id="167a8-421">要求</span><span class="sxs-lookup"><span data-stu-id="167a8-421">Requirement</span></span>| <span data-ttu-id="167a8-422">值</span><span class="sxs-lookup"><span data-stu-id="167a8-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="167a8-423">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="167a8-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167a8-424">1.0</span><span class="sxs-lookup"><span data-stu-id="167a8-424">1.0</span></span>|
|[<span data-ttu-id="167a8-425">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="167a8-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167a8-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="167a8-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="167a8-427">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="167a8-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167a8-428">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="167a8-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167a8-429">示例</span><span class="sxs-lookup"><span data-stu-id="167a8-429">Example</span></span>

<span data-ttu-id="167a8-430">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="167a8-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
