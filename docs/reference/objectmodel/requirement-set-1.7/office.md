 

# <a name="office"></a><span data-ttu-id="d700e-101">Office</span><span class="sxs-lookup"><span data-stu-id="d700e-101">Office</span></span>

<span data-ttu-id="d700e-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d700e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d700e-104">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-104">Requirements</span></span>

|<span data-ttu-id="d700e-105">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-105">Requirement</span></span>| <span data-ttu-id="d700e-106">值</span><span class="sxs-lookup"><span data-stu-id="d700e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d700e-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="d700e-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d700e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d700e-108">1.0</span></span>|
|[<span data-ttu-id="d700e-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d700e-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d700e-110">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="d700e-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d700e-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d700e-111">Members and methods</span></span>

| <span data-ttu-id="d700e-112">成员</span><span class="sxs-lookup"><span data-stu-id="d700e-112">Member</span></span> | <span data-ttu-id="d700e-113">类型</span><span class="sxs-lookup"><span data-stu-id="d700e-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d700e-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d700e-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d700e-115">成员</span><span class="sxs-lookup"><span data-stu-id="d700e-115">Member</span></span> |
| [<span data-ttu-id="d700e-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d700e-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d700e-117">成员</span><span class="sxs-lookup"><span data-stu-id="d700e-117">Member</span></span> |
| [<span data-ttu-id="d700e-118">EventType</span><span class="sxs-lookup"><span data-stu-id="d700e-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d700e-119">成员</span><span class="sxs-lookup"><span data-stu-id="d700e-119">Member</span></span> |
| [<span data-ttu-id="d700e-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d700e-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d700e-121">成员</span><span class="sxs-lookup"><span data-stu-id="d700e-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d700e-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="d700e-122">Namespaces</span></span>

<span data-ttu-id="d700e-123">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="d700e-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d700e-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="d700e-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d700e-125">成员</span><span class="sxs-lookup"><span data-stu-id="d700e-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d700e-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d700e-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="d700e-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d700e-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d700e-128">类型：</span><span class="sxs-lookup"><span data-stu-id="d700e-128">Type:</span></span>

*   <span data-ttu-id="d700e-129">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d700e-130">属性:</span><span class="sxs-lookup"><span data-stu-id="d700e-130">Properties:</span></span>

|<span data-ttu-id="d700e-131">名称</span><span class="sxs-lookup"><span data-stu-id="d700e-131">Name</span></span>| <span data-ttu-id="d700e-132">类型</span><span class="sxs-lookup"><span data-stu-id="d700e-132">Type</span></span>| <span data-ttu-id="d700e-133">说明</span><span class="sxs-lookup"><span data-stu-id="d700e-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d700e-134">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-134">String</span></span>|<span data-ttu-id="d700e-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d700e-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d700e-136">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-136">String</span></span>|<span data-ttu-id="d700e-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d700e-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d700e-138">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-138">Requirements</span></span>

|<span data-ttu-id="d700e-139">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-139">Requirement</span></span>| <span data-ttu-id="d700e-140">值</span><span class="sxs-lookup"><span data-stu-id="d700e-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="d700e-141">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="d700e-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d700e-142">1.0</span><span class="sxs-lookup"><span data-stu-id="d700e-142">1.0</span></span>|
|[<span data-ttu-id="d700e-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d700e-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d700e-144">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="d700e-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="d700e-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d700e-145">CoercionType :String</span></span>

<span data-ttu-id="d700e-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d700e-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d700e-147">类型：</span><span class="sxs-lookup"><span data-stu-id="d700e-147">Type:</span></span>

*   <span data-ttu-id="d700e-148">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d700e-149">属性:</span><span class="sxs-lookup"><span data-stu-id="d700e-149">Properties:</span></span>

|<span data-ttu-id="d700e-150">名称</span><span class="sxs-lookup"><span data-stu-id="d700e-150">Name</span></span>| <span data-ttu-id="d700e-151">类型</span><span class="sxs-lookup"><span data-stu-id="d700e-151">Type</span></span>| <span data-ttu-id="d700e-152">说明</span><span class="sxs-lookup"><span data-stu-id="d700e-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d700e-153">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-153">String</span></span>|<span data-ttu-id="d700e-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d700e-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d700e-155">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-155">String</span></span>|<span data-ttu-id="d700e-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d700e-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d700e-157">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-157">Requirements</span></span>

|<span data-ttu-id="d700e-158">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-158">Requirement</span></span>| <span data-ttu-id="d700e-159">值</span><span class="sxs-lookup"><span data-stu-id="d700e-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="d700e-160">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="d700e-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d700e-161">1.0</span><span class="sxs-lookup"><span data-stu-id="d700e-161">1.0</span></span>|
|[<span data-ttu-id="d700e-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d700e-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d700e-163">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="d700e-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="d700e-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="d700e-164">EventType :String</span></span>

<span data-ttu-id="d700e-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="d700e-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d700e-166">类型：</span><span class="sxs-lookup"><span data-stu-id="d700e-166">Type:</span></span>

*   <span data-ttu-id="d700e-167">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d700e-168">属性:</span><span class="sxs-lookup"><span data-stu-id="d700e-168">Properties:</span></span>

| <span data-ttu-id="d700e-169">名称</span><span class="sxs-lookup"><span data-stu-id="d700e-169">Name</span></span> | <span data-ttu-id="d700e-170">类型</span><span class="sxs-lookup"><span data-stu-id="d700e-170">Type</span></span> | <span data-ttu-id="d700e-171">说明</span><span class="sxs-lookup"><span data-stu-id="d700e-171">Description</span></span> | <span data-ttu-id="d700e-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="d700e-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="d700e-173">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-173">String</span></span> | <span data-ttu-id="d700e-174">所选约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="d700e-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="d700e-175">1.7</span><span class="sxs-lookup"><span data-stu-id="d700e-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="d700e-176">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-176">String</span></span> | <span data-ttu-id="d700e-177">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="d700e-177">The selected item has changed.</span></span> | <span data-ttu-id="d700e-178">1.5</span><span class="sxs-lookup"><span data-stu-id="d700e-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="d700e-179">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-179">String</span></span> | <span data-ttu-id="d700e-180">所选项或约会地点的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="d700e-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="d700e-181">1.7</span><span class="sxs-lookup"><span data-stu-id="d700e-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="d700e-182">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-182">String</span></span> | <span data-ttu-id="d700e-183">所选系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="d700e-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="d700e-184">1.7</span><span class="sxs-lookup"><span data-stu-id="d700e-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d700e-185">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-185">Requirements</span></span>

|<span data-ttu-id="d700e-186">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-186">Requirement</span></span>| <span data-ttu-id="d700e-187">值</span><span class="sxs-lookup"><span data-stu-id="d700e-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="d700e-188">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="d700e-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d700e-189">1.5</span><span class="sxs-lookup"><span data-stu-id="d700e-189">1.5</span></span> |
|[<span data-ttu-id="d700e-190">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d700e-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d700e-191">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="d700e-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="d700e-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d700e-192">SourceProperty :String</span></span>

<span data-ttu-id="d700e-193">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d700e-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d700e-194">类型：</span><span class="sxs-lookup"><span data-stu-id="d700e-194">Type:</span></span>

*   <span data-ttu-id="d700e-195">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d700e-196">属性:</span><span class="sxs-lookup"><span data-stu-id="d700e-196">Properties:</span></span>

|<span data-ttu-id="d700e-197">名称</span><span class="sxs-lookup"><span data-stu-id="d700e-197">Name</span></span>| <span data-ttu-id="d700e-198">类型</span><span class="sxs-lookup"><span data-stu-id="d700e-198">Type</span></span>| <span data-ttu-id="d700e-199">说明</span><span class="sxs-lookup"><span data-stu-id="d700e-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d700e-200">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-200">String</span></span>|<span data-ttu-id="d700e-201">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d700e-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d700e-202">字符串</span><span class="sxs-lookup"><span data-stu-id="d700e-202">String</span></span>|<span data-ttu-id="d700e-203">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d700e-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d700e-204">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-204">Requirements</span></span>

|<span data-ttu-id="d700e-205">要求</span><span class="sxs-lookup"><span data-stu-id="d700e-205">Requirement</span></span>| <span data-ttu-id="d700e-206">值</span><span class="sxs-lookup"><span data-stu-id="d700e-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="d700e-207">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="d700e-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d700e-208">1.0</span><span class="sxs-lookup"><span data-stu-id="d700e-208">1.0</span></span>|
|[<span data-ttu-id="d700e-209">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d700e-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d700e-210">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="d700e-210">Compose or read</span></span>|