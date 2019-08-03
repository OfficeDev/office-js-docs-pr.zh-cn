---
title: "\"Context.subname\"-\"邮箱-要求集 1.2\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: de7c48faf966f9b3f5d1bb76f69aa16810a5381f
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064366"
---
# <a name="mailbox"></a><span data-ttu-id="462a9-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="462a9-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="462a9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="462a9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="462a9-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="462a9-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="462a9-105">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-105">Requirements</span></span>

|<span data-ttu-id="462a9-106">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-106">Requirement</span></span>| <span data-ttu-id="462a9-107">值</span><span class="sxs-lookup"><span data-stu-id="462a9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-109">1.0</span></span>|
|[<span data-ttu-id="462a9-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-111">受限</span><span class="sxs-lookup"><span data-stu-id="462a9-111">Restricted</span></span>|
|[<span data-ttu-id="462a9-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="462a9-114">命名空间</span><span class="sxs-lookup"><span data-stu-id="462a9-114">Namespaces</span></span>

<span data-ttu-id="462a9-115">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="462a9-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="462a9-116">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="462a9-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="462a9-117">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="462a9-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="462a9-118">成员</span><span class="sxs-lookup"><span data-stu-id="462a9-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="462a9-119">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="462a9-119">ewsUrl: String</span></span>

<span data-ttu-id="462a9-120">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="462a9-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="462a9-121">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="462a9-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="462a9-122">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="462a9-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="462a9-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="462a9-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="462a9-125">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-125">Type</span></span>

*   <span data-ttu-id="462a9-126">String</span><span class="sxs-lookup"><span data-stu-id="462a9-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="462a9-127">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-127">Requirements</span></span>

|<span data-ttu-id="462a9-128">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-128">Requirement</span></span>| <span data-ttu-id="462a9-129">值</span><span class="sxs-lookup"><span data-stu-id="462a9-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-131">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-131">1.0</span></span>|
|[<span data-ttu-id="462a9-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-133">ReadItem</span></span>|
|[<span data-ttu-id="462a9-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-135">阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="462a9-136">方法</span><span class="sxs-lookup"><span data-stu-id="462a9-136">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="462a9-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="462a9-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="462a9-138">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="462a9-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="462a9-139">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="462a9-139">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="462a9-140">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="462a9-140">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="462a9-141">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="462a9-141">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="462a9-142">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="462a9-142">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="462a9-143">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="462a9-143">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-144">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-144">Parameters</span></span>

|<span data-ttu-id="462a9-145">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-145">Name</span></span>| <span data-ttu-id="462a9-146">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-146">Type</span></span>| <span data-ttu-id="462a9-147">描述</span><span class="sxs-lookup"><span data-stu-id="462a9-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="462a9-148">日期</span><span class="sxs-lookup"><span data-stu-id="462a9-148">Date</span></span>|<span data-ttu-id="462a9-149">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="462a9-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-150">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-150">Requirements</span></span>

|<span data-ttu-id="462a9-151">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-151">Requirement</span></span>| <span data-ttu-id="462a9-152">值</span><span class="sxs-lookup"><span data-stu-id="462a9-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-153">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-154">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-154">1.0</span></span>|
|[<span data-ttu-id="462a9-155">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-156">ReadItem</span></span>|
|[<span data-ttu-id="462a9-157">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-158">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-158">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462a9-159">返回：</span><span class="sxs-lookup"><span data-stu-id="462a9-159">Returns:</span></span>

<span data-ttu-id="462a9-160">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="462a9-160">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="462a9-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="462a9-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="462a9-162">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="462a9-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="462a9-163">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="462a9-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-164">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-164">Parameters</span></span>

|<span data-ttu-id="462a9-165">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-165">Name</span></span>| <span data-ttu-id="462a9-166">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-166">Type</span></span>| <span data-ttu-id="462a9-167">说明</span><span class="sxs-lookup"><span data-stu-id="462a9-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="462a9-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="462a9-168">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="462a9-169">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="462a9-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-170">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-170">Requirements</span></span>

|<span data-ttu-id="462a9-171">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-171">Requirement</span></span>| <span data-ttu-id="462a9-172">值</span><span class="sxs-lookup"><span data-stu-id="462a9-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-173">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-174">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-174">1.0</span></span>|
|[<span data-ttu-id="462a9-175">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-176">ReadItem</span></span>|
|[<span data-ttu-id="462a9-177">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-178">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-178">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462a9-179">返回：</span><span class="sxs-lookup"><span data-stu-id="462a9-179">Returns:</span></span>

<span data-ttu-id="462a9-180">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="462a9-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="462a9-181">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="462a9-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="462a9-182">日期</span><span class="sxs-lookup"><span data-stu-id="462a9-182">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="462a9-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="462a9-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="462a9-184">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="462a9-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="462a9-185">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="462a9-185">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="462a9-186">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="462a9-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="462a9-187">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="462a9-187">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="462a9-188">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="462a9-188">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="462a9-189">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="462a9-189">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="462a9-190">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="462a9-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-191">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-191">Parameters</span></span>

|<span data-ttu-id="462a9-192">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-192">Name</span></span>| <span data-ttu-id="462a9-193">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-193">Type</span></span>| <span data-ttu-id="462a9-194">说明</span><span class="sxs-lookup"><span data-stu-id="462a9-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="462a9-195">字符串</span><span class="sxs-lookup"><span data-stu-id="462a9-195">String</span></span>|<span data-ttu-id="462a9-196">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="462a9-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-197">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-197">Requirements</span></span>

|<span data-ttu-id="462a9-198">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-198">Requirement</span></span>| <span data-ttu-id="462a9-199">值</span><span class="sxs-lookup"><span data-stu-id="462a9-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-201">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-201">1.0</span></span>|
|[<span data-ttu-id="462a9-202">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-203">ReadItem</span></span>|
|[<span data-ttu-id="462a9-204">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-205">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-205">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462a9-206">示例</span><span class="sxs-lookup"><span data-stu-id="462a9-206">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="462a9-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="462a9-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="462a9-208">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="462a9-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="462a9-209">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="462a9-209">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="462a9-210">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="462a9-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="462a9-211">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="462a9-211">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="462a9-212">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="462a9-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="462a9-p106">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="462a9-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-215">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-215">Parameters</span></span>

|<span data-ttu-id="462a9-216">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-216">Name</span></span>| <span data-ttu-id="462a9-217">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-217">Type</span></span>| <span data-ttu-id="462a9-218">说明</span><span class="sxs-lookup"><span data-stu-id="462a9-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="462a9-219">String</span><span class="sxs-lookup"><span data-stu-id="462a9-219">String</span></span>|<span data-ttu-id="462a9-220">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="462a9-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-221">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-221">Requirements</span></span>

|<span data-ttu-id="462a9-222">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-222">Requirement</span></span>| <span data-ttu-id="462a9-223">值</span><span class="sxs-lookup"><span data-stu-id="462a9-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-225">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-225">1.0</span></span>|
|[<span data-ttu-id="462a9-226">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-226">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-227">ReadItem</span></span>|
|[<span data-ttu-id="462a9-228">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-228">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-229">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-229">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462a9-230">示例</span><span class="sxs-lookup"><span data-stu-id="462a9-230">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="462a9-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="462a9-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="462a9-232">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="462a9-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="462a9-233">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="462a9-233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="462a9-p107">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="462a9-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="462a9-236">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="462a9-236">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="462a9-237">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="462a9-237">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="462a9-238">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="462a9-238">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="462a9-p109">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="462a9-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="462a9-241">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="462a9-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-242">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-242">Parameters</span></span>

|<span data-ttu-id="462a9-243">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-243">Name</span></span>| <span data-ttu-id="462a9-244">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-244">Type</span></span>| <span data-ttu-id="462a9-245">描述</span><span class="sxs-lookup"><span data-stu-id="462a9-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="462a9-246">对象</span><span class="sxs-lookup"><span data-stu-id="462a9-246">Object</span></span> | <span data-ttu-id="462a9-247">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="462a9-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="462a9-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="462a9-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="462a9-p110">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="462a9-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="462a9-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="462a9-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="462a9-p111">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="462a9-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="462a9-254">日期</span><span class="sxs-lookup"><span data-stu-id="462a9-254">Date</span></span> | <span data-ttu-id="462a9-255">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="462a9-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="462a9-256">Date</span><span class="sxs-lookup"><span data-stu-id="462a9-256">Date</span></span> | <span data-ttu-id="462a9-257">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="462a9-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="462a9-258">String</span><span class="sxs-lookup"><span data-stu-id="462a9-258">String</span></span> | <span data-ttu-id="462a9-p112">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="462a9-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="462a9-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="462a9-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="462a9-p113">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="462a9-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="462a9-264">String</span><span class="sxs-lookup"><span data-stu-id="462a9-264">String</span></span> | <span data-ttu-id="462a9-p114">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="462a9-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="462a9-267">字符串</span><span class="sxs-lookup"><span data-stu-id="462a9-267">String</span></span> | <span data-ttu-id="462a9-p115">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="462a9-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462a9-270">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-270">Requirements</span></span>

|<span data-ttu-id="462a9-271">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-271">Requirement</span></span>| <span data-ttu-id="462a9-272">值</span><span class="sxs-lookup"><span data-stu-id="462a9-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-274">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-274">1.0</span></span>|
|[<span data-ttu-id="462a9-275">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-276">ReadItem</span></span>|
|[<span data-ttu-id="462a9-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-278">阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462a9-279">示例</span><span class="sxs-lookup"><span data-stu-id="462a9-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="462a9-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="462a9-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="462a9-281">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="462a9-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="462a9-p116">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="462a9-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="462a9-p117">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="462a9-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="462a9-287">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="462a9-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-288">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-288">Parameters</span></span>

|<span data-ttu-id="462a9-289">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-289">Name</span></span>| <span data-ttu-id="462a9-290">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-290">Type</span></span>| <span data-ttu-id="462a9-291">属性</span><span class="sxs-lookup"><span data-stu-id="462a9-291">Attributes</span></span>| <span data-ttu-id="462a9-292">说明</span><span class="sxs-lookup"><span data-stu-id="462a9-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="462a9-293">函数</span><span class="sxs-lookup"><span data-stu-id="462a9-293">function</span></span>||<span data-ttu-id="462a9-294">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="462a9-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="462a9-295">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="462a9-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="462a9-296">对象</span><span class="sxs-lookup"><span data-stu-id="462a9-296">Object</span></span>| <span data-ttu-id="462a9-297">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="462a9-297">&lt;optional&gt;</span></span>|<span data-ttu-id="462a9-298">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="462a9-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-299">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-299">Requirements</span></span>

|<span data-ttu-id="462a9-300">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-300">Requirement</span></span>| <span data-ttu-id="462a9-301">值</span><span class="sxs-lookup"><span data-stu-id="462a9-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-303">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-303">1.0</span></span>|
|[<span data-ttu-id="462a9-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-305">ReadItem</span></span>|
|[<span data-ttu-id="462a9-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-307">阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462a9-308">示例</span><span class="sxs-lookup"><span data-stu-id="462a9-308">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="462a9-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="462a9-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="462a9-310">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="462a9-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="462a9-311">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="462a9-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-312">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-312">Parameters</span></span>

|<span data-ttu-id="462a9-313">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-313">Name</span></span>| <span data-ttu-id="462a9-314">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-314">Type</span></span>| <span data-ttu-id="462a9-315">属性</span><span class="sxs-lookup"><span data-stu-id="462a9-315">Attributes</span></span>| <span data-ttu-id="462a9-316">说明</span><span class="sxs-lookup"><span data-stu-id="462a9-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="462a9-317">函数</span><span class="sxs-lookup"><span data-stu-id="462a9-317">function</span></span>||<span data-ttu-id="462a9-318">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="462a9-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="462a9-319">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="462a9-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="462a9-320">对象</span><span class="sxs-lookup"><span data-stu-id="462a9-320">Object</span></span>| <span data-ttu-id="462a9-321">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="462a9-321">&lt;optional&gt;</span></span>|<span data-ttu-id="462a9-322">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="462a9-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-323">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-323">Requirements</span></span>

|<span data-ttu-id="462a9-324">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-324">Requirement</span></span>| <span data-ttu-id="462a9-325">值</span><span class="sxs-lookup"><span data-stu-id="462a9-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-326">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-327">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-327">1.0</span></span>|
|[<span data-ttu-id="462a9-328">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462a9-329">ReadItem</span></span>|
|[<span data-ttu-id="462a9-330">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-331">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-331">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462a9-332">示例</span><span class="sxs-lookup"><span data-stu-id="462a9-332">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="462a9-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="462a9-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="462a9-334">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="462a9-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="462a9-335">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="462a9-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="462a9-336">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="462a9-336">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="462a9-337">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="462a9-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="462a9-338">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="462a9-338">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="462a9-339">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="462a9-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="462a9-340">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="462a9-340">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="462a9-341">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="462a9-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="462a9-342">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="462a9-342">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="462a9-p119">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="462a9-p119">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="462a9-345">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="462a9-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="462a9-346">版本差异</span><span class="sxs-lookup"><span data-stu-id="462a9-346">Version differences</span></span>

<span data-ttu-id="462a9-347">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="462a9-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="462a9-p120">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="462a9-p120">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462a9-351">参数</span><span class="sxs-lookup"><span data-stu-id="462a9-351">Parameters</span></span>

|<span data-ttu-id="462a9-352">名称</span><span class="sxs-lookup"><span data-stu-id="462a9-352">Name</span></span>| <span data-ttu-id="462a9-353">类型</span><span class="sxs-lookup"><span data-stu-id="462a9-353">Type</span></span>| <span data-ttu-id="462a9-354">属性</span><span class="sxs-lookup"><span data-stu-id="462a9-354">Attributes</span></span>| <span data-ttu-id="462a9-355">说明</span><span class="sxs-lookup"><span data-stu-id="462a9-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="462a9-356">字符串</span><span class="sxs-lookup"><span data-stu-id="462a9-356">String</span></span>||<span data-ttu-id="462a9-357">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="462a9-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="462a9-358">函数</span><span class="sxs-lookup"><span data-stu-id="462a9-358">function</span></span>||<span data-ttu-id="462a9-359">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="462a9-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="462a9-360">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="462a9-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="462a9-361">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="462a9-361">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="462a9-362">对象</span><span class="sxs-lookup"><span data-stu-id="462a9-362">Object</span></span>| <span data-ttu-id="462a9-363">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="462a9-363">&lt;optional&gt;</span></span>|<span data-ttu-id="462a9-364">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="462a9-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462a9-365">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-365">Requirements</span></span>

|<span data-ttu-id="462a9-366">要求</span><span class="sxs-lookup"><span data-stu-id="462a9-366">Requirement</span></span>| <span data-ttu-id="462a9-367">值</span><span class="sxs-lookup"><span data-stu-id="462a9-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="462a9-368">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="462a9-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462a9-369">1.0</span><span class="sxs-lookup"><span data-stu-id="462a9-369">1.0</span></span>|
|[<span data-ttu-id="462a9-370">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="462a9-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462a9-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="462a9-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="462a9-372">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="462a9-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="462a9-373">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="462a9-373">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462a9-374">示例</span><span class="sxs-lookup"><span data-stu-id="462a9-374">Example</span></span>

<span data-ttu-id="462a9-375">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="462a9-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
