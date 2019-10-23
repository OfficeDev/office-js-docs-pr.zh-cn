---
title: "\"Context.subname\"-\"邮箱-要求集 1.2\""
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 542e8c9899c2d4a3c5b4546c3d5a73ba0d3c3a7e
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626997"
---
# <a name="mailbox"></a><span data-ttu-id="5c4bd-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="5c4bd-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="5c4bd-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="5c4bd-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="5c4bd-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c4bd-105">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-105">Requirements</span></span>

|<span data-ttu-id="5c4bd-106">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-106">Requirement</span></span>| <span data-ttu-id="5c4bd-107">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-109">1.0</span></span>|
|[<span data-ttu-id="5c4bd-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-111">受限</span><span class="sxs-lookup"><span data-stu-id="5c4bd-111">Restricted</span></span>|
|[<span data-ttu-id="5c4bd-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5c4bd-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-114">Members and methods</span></span>

| <span data-ttu-id="5c4bd-115">成员</span><span class="sxs-lookup"><span data-stu-id="5c4bd-115">Member</span></span> | <span data-ttu-id="5c4bd-116">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5c4bd-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="5c4bd-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="5c4bd-118">成员</span><span class="sxs-lookup"><span data-stu-id="5c4bd-118">Member</span></span> |
| [<span data-ttu-id="5c4bd-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="5c4bd-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="5c4bd-120">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-120">Method</span></span> |
| [<span data-ttu-id="5c4bd-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="5c4bd-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="5c4bd-122">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-122">Method</span></span> |
| [<span data-ttu-id="5c4bd-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="5c4bd-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="5c4bd-124">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-124">Method</span></span> |
| [<span data-ttu-id="5c4bd-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="5c4bd-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="5c4bd-126">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-126">Method</span></span> |
| [<span data-ttu-id="5c4bd-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="5c4bd-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="5c4bd-128">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-128">Method</span></span> |
| [<span data-ttu-id="5c4bd-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="5c4bd-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="5c4bd-130">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-130">Method</span></span> |
| [<span data-ttu-id="5c4bd-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="5c4bd-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="5c4bd-132">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-132">Method</span></span> |
| [<span data-ttu-id="5c4bd-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="5c4bd-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="5c4bd-134">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5c4bd-135">命名空间</span><span class="sxs-lookup"><span data-stu-id="5c4bd-135">Namespaces</span></span>

<span data-ttu-id="5c4bd-136">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="5c4bd-137">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="5c4bd-138">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="5c4bd-139">Members</span><span class="sxs-lookup"><span data-stu-id="5c4bd-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="5c4bd-140">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="5c4bd-140">ewsUrl: String</span></span>

<span data-ttu-id="5c4bd-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4bd-143">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5c4bd-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="5c4bd-146">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-146">Type</span></span>

*   <span data-ttu-id="5c4bd-147">String</span><span class="sxs-lookup"><span data-stu-id="5c4bd-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c4bd-148">Requirements</span><span class="sxs-lookup"><span data-stu-id="5c4bd-148">Requirements</span></span>

|<span data-ttu-id="5c4bd-149">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-149">Requirement</span></span>| <span data-ttu-id="5c4bd-150">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-151">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-152">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-152">1.0</span></span>|
|[<span data-ttu-id="5c4bd-153">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-154">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-155">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-156">阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="5c4bd-157">方法</span><span class="sxs-lookup"><span data-stu-id="5c4bd-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="5c4bd-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="5c4bd-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="5c4bd-159">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="5c4bd-p103">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p103">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="5c4bd-p104">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p104">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-165">Parameters</span><span class="sxs-lookup"><span data-stu-id="5c4bd-165">Parameters</span></span>

|<span data-ttu-id="5c4bd-166">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-166">Name</span></span>| <span data-ttu-id="5c4bd-167">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-167">Type</span></span>| <span data-ttu-id="5c4bd-168">描述</span><span class="sxs-lookup"><span data-stu-id="5c4bd-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="5c4bd-169">日期</span><span class="sxs-lookup"><span data-stu-id="5c4bd-169">Date</span></span>|<span data-ttu-id="5c4bd-170">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="5c4bd-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-171">Requirements</span><span class="sxs-lookup"><span data-stu-id="5c4bd-171">Requirements</span></span>

|<span data-ttu-id="5c4bd-172">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-172">Requirement</span></span>| <span data-ttu-id="5c4bd-173">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-175">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-175">1.0</span></span>|
|[<span data-ttu-id="5c4bd-176">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-177">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-178">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-179">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5c4bd-180">返回：</span><span class="sxs-lookup"><span data-stu-id="5c4bd-180">Returns:</span></span>

<span data-ttu-id="5c4bd-181">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5c4bd-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="5c4bd-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="5c4bd-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="5c4bd-183">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="5c4bd-184">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-185">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-185">Parameters</span></span>

|<span data-ttu-id="5c4bd-186">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-186">Name</span></span>| <span data-ttu-id="5c4bd-187">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-187">Type</span></span>| <span data-ttu-id="5c4bd-188">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="5c4bd-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="5c4bd-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="5c4bd-190">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-191">Requirements</span><span class="sxs-lookup"><span data-stu-id="5c4bd-191">Requirements</span></span>

|<span data-ttu-id="5c4bd-192">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-192">Requirement</span></span>| <span data-ttu-id="5c4bd-193">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-195">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-195">1.0</span></span>|
|[<span data-ttu-id="5c4bd-196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-197">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5c4bd-200">返回：</span><span class="sxs-lookup"><span data-stu-id="5c4bd-200">Returns:</span></span>

<span data-ttu-id="5c4bd-201">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="5c4bd-202">键入：日期</span><span class="sxs-lookup"><span data-stu-id="5c4bd-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="5c4bd-203">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-203">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="5c4bd-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="5c4bd-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="5c4bd-205">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4bd-206">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5c4bd-207">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="5c4bd-p105">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p105">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="5c4bd-210">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="5c4bd-211">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-212">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-212">Parameters</span></span>

|<span data-ttu-id="5c4bd-213">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-213">Name</span></span>| <span data-ttu-id="5c4bd-214">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-214">Type</span></span>| <span data-ttu-id="5c4bd-215">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="5c4bd-216">字符串</span><span class="sxs-lookup"><span data-stu-id="5c4bd-216">String</span></span>|<span data-ttu-id="5c4bd-217">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-218">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-218">Requirements</span></span>

|<span data-ttu-id="5c4bd-219">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-219">Requirement</span></span>| <span data-ttu-id="5c4bd-220">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-221">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-222">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-222">1.0</span></span>|
|[<span data-ttu-id="5c4bd-223">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-224">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c4bd-227">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="5c4bd-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="5c4bd-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="5c4bd-229">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4bd-230">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5c4bd-231">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="5c4bd-232">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="5c4bd-233">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="5c4bd-p106">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-236">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-236">Parameters</span></span>

|<span data-ttu-id="5c4bd-237">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-237">Name</span></span>| <span data-ttu-id="5c4bd-238">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-238">Type</span></span>| <span data-ttu-id="5c4bd-239">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="5c4bd-240">String</span><span class="sxs-lookup"><span data-stu-id="5c4bd-240">String</span></span>|<span data-ttu-id="5c4bd-241">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-242">Requirements</span><span class="sxs-lookup"><span data-stu-id="5c4bd-242">Requirements</span></span>

|<span data-ttu-id="5c4bd-243">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-243">Requirement</span></span>| <span data-ttu-id="5c4bd-244">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-245">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-246">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-246">1.0</span></span>|
|[<span data-ttu-id="5c4bd-247">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-248">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-249">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-250">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c4bd-251">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="5c4bd-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="5c4bd-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="5c4bd-253">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4bd-254">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5c4bd-p107">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="5c4bd-p108">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p108">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="5c4bd-p109">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="5c4bd-262">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-263">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-263">Parameters</span></span>

|<span data-ttu-id="5c4bd-264">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-264">Name</span></span>| <span data-ttu-id="5c4bd-265">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-265">Type</span></span>| <span data-ttu-id="5c4bd-266">描述</span><span class="sxs-lookup"><span data-stu-id="5c4bd-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="5c4bd-267">对象</span><span class="sxs-lookup"><span data-stu-id="5c4bd-267">Object</span></span> | <span data-ttu-id="5c4bd-268">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="5c4bd-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="5c4bd-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="5c4bd-p110">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="5c4bd-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="5c4bd-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="5c4bd-p111">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="5c4bd-275">日期</span><span class="sxs-lookup"><span data-stu-id="5c4bd-275">Date</span></span> | <span data-ttu-id="5c4bd-276">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="5c4bd-277">Date</span><span class="sxs-lookup"><span data-stu-id="5c4bd-277">Date</span></span> | <span data-ttu-id="5c4bd-278">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="5c4bd-279">String</span><span class="sxs-lookup"><span data-stu-id="5c4bd-279">String</span></span> | <span data-ttu-id="5c4bd-p112">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="5c4bd-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="5c4bd-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="5c4bd-p113">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="5c4bd-285">String</span><span class="sxs-lookup"><span data-stu-id="5c4bd-285">String</span></span> | <span data-ttu-id="5c4bd-p114">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="5c4bd-288">字符串</span><span class="sxs-lookup"><span data-stu-id="5c4bd-288">String</span></span> | <span data-ttu-id="5c4bd-p115">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5c4bd-291">Requirements</span><span class="sxs-lookup"><span data-stu-id="5c4bd-291">Requirements</span></span>

|<span data-ttu-id="5c4bd-292">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-292">Requirement</span></span>| <span data-ttu-id="5c4bd-293">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-294">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-295">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-295">1.0</span></span>|
|[<span data-ttu-id="5c4bd-296">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-297">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-298">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-299">阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c4bd-300">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-300">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="5c4bd-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5c4bd-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="5c4bd-302">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="5c4bd-p116">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="5c4bd-305">您可以将令牌和附件标识符或项目标识符同时传递给第三方系统。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-305">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="5c4bd-306">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-306">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="5c4bd-307">例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-307">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="5c4bd-308">调用`getCallbackTokenAsync`方法需要**ReadItem**的最低权限级别。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-308">Calling the `getCallbackTokenAsync` method requires a minimum permission level of **ReadItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-309">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-309">Parameters</span></span>

|<span data-ttu-id="5c4bd-310">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-310">Name</span></span>| <span data-ttu-id="5c4bd-311">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-311">Type</span></span>| <span data-ttu-id="5c4bd-312">属性</span><span class="sxs-lookup"><span data-stu-id="5c4bd-312">Attributes</span></span>| <span data-ttu-id="5c4bd-313">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5c4bd-314">函数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-314">function</span></span>||<span data-ttu-id="5c4bd-315">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5c4bd-316">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="5c4bd-317">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="5c4bd-318">对象</span><span class="sxs-lookup"><span data-stu-id="5c4bd-318">Object</span></span>| <span data-ttu-id="5c4bd-319">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="5c4bd-319">&lt;optional&gt;</span></span>|<span data-ttu-id="5c4bd-320">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5c4bd-321">错误</span><span class="sxs-lookup"><span data-stu-id="5c4bd-321">Errors</span></span>

|<span data-ttu-id="5c4bd-322">错误代码</span><span class="sxs-lookup"><span data-stu-id="5c4bd-322">Error code</span></span>|<span data-ttu-id="5c4bd-323">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="5c4bd-324">请求失败。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-324">The request has failed.</span></span> <span data-ttu-id="5c4bd-325">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="5c4bd-326">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-326">The Exchange server returned an error.</span></span> <span data-ttu-id="5c4bd-327">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="5c4bd-328">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="5c4bd-329">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-330">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-330">Requirements</span></span>

|<span data-ttu-id="5c4bd-331">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-331">Requirement</span></span>| <span data-ttu-id="5c4bd-332">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-333">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-334">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-334">1.0</span></span>|
|[<span data-ttu-id="5c4bd-335">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-336">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-337">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-338">阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c4bd-339">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-339">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="5c4bd-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5c4bd-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="5c4bd-341">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="5c4bd-342">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-343">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-343">Parameters</span></span>

|<span data-ttu-id="5c4bd-344">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-344">Name</span></span>| <span data-ttu-id="5c4bd-345">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-345">Type</span></span>| <span data-ttu-id="5c4bd-346">属性</span><span class="sxs-lookup"><span data-stu-id="5c4bd-346">Attributes</span></span>| <span data-ttu-id="5c4bd-347">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5c4bd-348">函数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-348">function</span></span>||<span data-ttu-id="5c4bd-349">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5c4bd-350">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="5c4bd-351">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="5c4bd-352">对象</span><span class="sxs-lookup"><span data-stu-id="5c4bd-352">Object</span></span>| <span data-ttu-id="5c4bd-353">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="5c4bd-353">&lt;optional&gt;</span></span>|<span data-ttu-id="5c4bd-354">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5c4bd-355">错误</span><span class="sxs-lookup"><span data-stu-id="5c4bd-355">Errors</span></span>

|<span data-ttu-id="5c4bd-356">错误代码</span><span class="sxs-lookup"><span data-stu-id="5c4bd-356">Error code</span></span>|<span data-ttu-id="5c4bd-357">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="5c4bd-358">请求失败。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-358">The request has failed.</span></span> <span data-ttu-id="5c4bd-359">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="5c4bd-360">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-360">The Exchange server returned an error.</span></span> <span data-ttu-id="5c4bd-361">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="5c4bd-362">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="5c4bd-363">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-364">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-364">Requirements</span></span>

|<span data-ttu-id="5c4bd-365">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-365">Requirement</span></span>| <span data-ttu-id="5c4bd-366">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-367">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-368">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-368">1.0</span></span>|
|[<span data-ttu-id="5c4bd-369">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c4bd-370">ReadItem</span></span>|
|[<span data-ttu-id="5c4bd-371">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-372">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c4bd-373">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-373">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="5c4bd-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5c4bd-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="5c4bd-375">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4bd-376">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="5c4bd-377">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="5c4bd-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="5c4bd-378">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="5c4bd-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="5c4bd-379">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="5c4bd-380">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-380">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="5c4bd-381">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-381">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="5c4bd-382">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="5c4bd-383">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="5c4bd-p125">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="5c4bd-386">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="5c4bd-387">版本差异</span><span class="sxs-lookup"><span data-stu-id="5c4bd-387">Version differences</span></span>

<span data-ttu-id="5c4bd-388">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="5c4bd-p126">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5c4bd-392">参数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-392">Parameters</span></span>

|<span data-ttu-id="5c4bd-393">名称</span><span class="sxs-lookup"><span data-stu-id="5c4bd-393">Name</span></span>| <span data-ttu-id="5c4bd-394">类型</span><span class="sxs-lookup"><span data-stu-id="5c4bd-394">Type</span></span>| <span data-ttu-id="5c4bd-395">属性</span><span class="sxs-lookup"><span data-stu-id="5c4bd-395">Attributes</span></span>| <span data-ttu-id="5c4bd-396">说明</span><span class="sxs-lookup"><span data-stu-id="5c4bd-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5c4bd-397">字符串</span><span class="sxs-lookup"><span data-stu-id="5c4bd-397">String</span></span>||<span data-ttu-id="5c4bd-398">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="5c4bd-399">函数</span><span class="sxs-lookup"><span data-stu-id="5c4bd-399">function</span></span>||<span data-ttu-id="5c4bd-400">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5c4bd-401">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-401">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="5c4bd-402">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-402">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="5c4bd-403">对象</span><span class="sxs-lookup"><span data-stu-id="5c4bd-403">Object</span></span>| <span data-ttu-id="5c4bd-404">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="5c4bd-404">&lt;optional&gt;</span></span>|<span data-ttu-id="5c4bd-405">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5c4bd-406">Requirements</span><span class="sxs-lookup"><span data-stu-id="5c4bd-406">Requirements</span></span>

|<span data-ttu-id="5c4bd-407">要求</span><span class="sxs-lookup"><span data-stu-id="5c4bd-407">Requirement</span></span>| <span data-ttu-id="5c4bd-408">值</span><span class="sxs-lookup"><span data-stu-id="5c4bd-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c4bd-409">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c4bd-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c4bd-410">1.0</span><span class="sxs-lookup"><span data-stu-id="5c4bd-410">1.0</span></span>|
|[<span data-ttu-id="5c4bd-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c4bd-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c4bd-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="5c4bd-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="5c4bd-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c4bd-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c4bd-414">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c4bd-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c4bd-415">示例</span><span class="sxs-lookup"><span data-stu-id="5c4bd-415">Example</span></span>

<span data-ttu-id="5c4bd-416">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="5c4bd-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
