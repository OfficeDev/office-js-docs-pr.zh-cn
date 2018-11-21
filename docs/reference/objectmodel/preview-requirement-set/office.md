 

# <a name="office"></a><span data-ttu-id="a46ad-101">Office</span><span class="sxs-lookup"><span data-stu-id="a46ad-101">Office</span></span>

<span data-ttu-id="a46ad-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="a46ad-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a46ad-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="a46ad-104">Requirements</span></span>

|<span data-ttu-id="a46ad-105">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-105">Requirement</span></span>| <span data-ttu-id="a46ad-106">值</span><span class="sxs-lookup"><span data-stu-id="a46ad-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a46ad-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a46ad-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a46ad-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a46ad-108">1.0</span></span>|
|[<span data-ttu-id="a46ad-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a46ad-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a46ad-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a46ad-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a46ad-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="a46ad-111">Members and methods</span></span>

| <span data-ttu-id="a46ad-112">成员</span><span class="sxs-lookup"><span data-stu-id="a46ad-112">Member</span></span> | <span data-ttu-id="a46ad-113">类型</span><span class="sxs-lookup"><span data-stu-id="a46ad-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a46ad-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a46ad-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a46ad-115">成员</span><span class="sxs-lookup"><span data-stu-id="a46ad-115">Member</span></span> |
| [<span data-ttu-id="a46ad-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a46ad-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a46ad-117">成员</span><span class="sxs-lookup"><span data-stu-id="a46ad-117">Member</span></span> |
| [<span data-ttu-id="a46ad-118">EventType</span><span class="sxs-lookup"><span data-stu-id="a46ad-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a46ad-119">成员</span><span class="sxs-lookup"><span data-stu-id="a46ad-119">Member</span></span> |
| [<span data-ttu-id="a46ad-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a46ad-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a46ad-121">成员</span><span class="sxs-lookup"><span data-stu-id="a46ad-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a46ad-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="a46ad-122">Namespaces</span></span>

<span data-ttu-id="a46ad-123">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="a46ad-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a46ad-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="a46ad-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a46ad-125">成员</span><span class="sxs-lookup"><span data-stu-id="a46ad-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a46ad-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a46ad-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="a46ad-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="a46ad-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a46ad-128">类型：</span><span class="sxs-lookup"><span data-stu-id="a46ad-128">Type:</span></span>

*   <span data-ttu-id="a46ad-129">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a46ad-130">属性：</span><span class="sxs-lookup"><span data-stu-id="a46ad-130">Properties:</span></span>

|<span data-ttu-id="a46ad-131">名称</span><span class="sxs-lookup"><span data-stu-id="a46ad-131">Name</span></span>| <span data-ttu-id="a46ad-132">类型</span><span class="sxs-lookup"><span data-stu-id="a46ad-132">Type</span></span>| <span data-ttu-id="a46ad-133">描述</span><span class="sxs-lookup"><span data-stu-id="a46ad-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a46ad-134">String</span><span class="sxs-lookup"><span data-stu-id="a46ad-134">String</span></span>|<span data-ttu-id="a46ad-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="a46ad-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a46ad-136">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-136">String</span></span>|<span data-ttu-id="a46ad-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="a46ad-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a46ad-138">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-138">Requirements</span></span>

|<span data-ttu-id="a46ad-139">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-139">Requirement</span></span>| <span data-ttu-id="a46ad-140">值</span><span class="sxs-lookup"><span data-stu-id="a46ad-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="a46ad-141">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a46ad-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a46ad-142">1.0</span><span class="sxs-lookup"><span data-stu-id="a46ad-142">1.0</span></span>|
|[<span data-ttu-id="a46ad-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a46ad-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a46ad-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a46ad-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="a46ad-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a46ad-145">CoercionType :String</span></span>

<span data-ttu-id="a46ad-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="a46ad-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a46ad-147">类型：</span><span class="sxs-lookup"><span data-stu-id="a46ad-147">Type:</span></span>

*   <span data-ttu-id="a46ad-148">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a46ad-149">属性：</span><span class="sxs-lookup"><span data-stu-id="a46ad-149">Properties:</span></span>

|<span data-ttu-id="a46ad-150">名称</span><span class="sxs-lookup"><span data-stu-id="a46ad-150">Name</span></span>| <span data-ttu-id="a46ad-151">类型</span><span class="sxs-lookup"><span data-stu-id="a46ad-151">Type</span></span>| <span data-ttu-id="a46ad-152">描述</span><span class="sxs-lookup"><span data-stu-id="a46ad-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a46ad-153">String</span><span class="sxs-lookup"><span data-stu-id="a46ad-153">String</span></span>|<span data-ttu-id="a46ad-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="a46ad-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a46ad-155">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-155">String</span></span>|<span data-ttu-id="a46ad-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="a46ad-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a46ad-157">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-157">Requirements</span></span>

|<span data-ttu-id="a46ad-158">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-158">Requirement</span></span>| <span data-ttu-id="a46ad-159">值</span><span class="sxs-lookup"><span data-stu-id="a46ad-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="a46ad-160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a46ad-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a46ad-161">1.0</span><span class="sxs-lookup"><span data-stu-id="a46ad-161">1.0</span></span>|
|[<span data-ttu-id="a46ad-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a46ad-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a46ad-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a46ad-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="a46ad-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="a46ad-164">EventType :String</span></span>

<span data-ttu-id="a46ad-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="a46ad-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a46ad-166">类型：</span><span class="sxs-lookup"><span data-stu-id="a46ad-166">Type:</span></span>

*   <span data-ttu-id="a46ad-167">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a46ad-168">属性：</span><span class="sxs-lookup"><span data-stu-id="a46ad-168">Properties:</span></span>

| <span data-ttu-id="a46ad-169">名称</span><span class="sxs-lookup"><span data-stu-id="a46ad-169">Name</span></span> | <span data-ttu-id="a46ad-170">类型</span><span class="sxs-lookup"><span data-stu-id="a46ad-170">Type</span></span> | <span data-ttu-id="a46ad-171">描述</span><span class="sxs-lookup"><span data-stu-id="a46ad-171">Description</span></span> | <span data-ttu-id="a46ad-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="a46ad-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="a46ad-173">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-173">String</span></span> | <span data-ttu-id="a46ad-174">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="a46ad-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="a46ad-175">1.7</span><span class="sxs-lookup"><span data-stu-id="a46ad-175">-17</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="a46ad-176">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-176">String</span></span> | <span data-ttu-id="a46ad-177">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="a46ad-177">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="a46ad-178">预览</span><span class="sxs-lookup"><span data-stu-id="a46ad-178">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="a46ad-179">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-179">String</span></span> | <span data-ttu-id="a46ad-180">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="a46ad-180">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a46ad-181">1.5</span><span class="sxs-lookup"><span data-stu-id="a46ad-181">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="a46ad-182">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-182">String</span></span> | <span data-ttu-id="a46ad-183">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="a46ad-183">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="a46ad-184">预览</span><span class="sxs-lookup"><span data-stu-id="a46ad-184">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="a46ad-185">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-185">String</span></span> | <span data-ttu-id="a46ad-186">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="a46ad-186">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="a46ad-187">1.7</span><span class="sxs-lookup"><span data-stu-id="a46ad-187">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="a46ad-188">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-188">String</span></span> | <span data-ttu-id="a46ad-189">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="a46ad-189">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="a46ad-190">1.7</span><span class="sxs-lookup"><span data-stu-id="a46ad-190">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a46ad-191">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-191">Requirements</span></span>

|<span data-ttu-id="a46ad-192">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-192">Requirement</span></span>| <span data-ttu-id="a46ad-193">值</span><span class="sxs-lookup"><span data-stu-id="a46ad-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="a46ad-194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a46ad-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a46ad-195">1.5</span><span class="sxs-lookup"><span data-stu-id="a46ad-195">1.5</span></span> |
|[<span data-ttu-id="a46ad-196">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a46ad-196">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a46ad-197">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a46ad-197">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="a46ad-198">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a46ad-198">SourceProperty :String</span></span>

<span data-ttu-id="a46ad-199">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="a46ad-199">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a46ad-200">类型：</span><span class="sxs-lookup"><span data-stu-id="a46ad-200">Type:</span></span>

*   <span data-ttu-id="a46ad-201">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-201">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a46ad-202">属性：</span><span class="sxs-lookup"><span data-stu-id="a46ad-202">Properties:</span></span>

|<span data-ttu-id="a46ad-203">名称</span><span class="sxs-lookup"><span data-stu-id="a46ad-203">Name</span></span>| <span data-ttu-id="a46ad-204">类型</span><span class="sxs-lookup"><span data-stu-id="a46ad-204">Type</span></span>| <span data-ttu-id="a46ad-205">描述</span><span class="sxs-lookup"><span data-stu-id="a46ad-205">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a46ad-206">字符串</span><span class="sxs-lookup"><span data-stu-id="a46ad-206">String</span></span>|<span data-ttu-id="a46ad-207">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="a46ad-207">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a46ad-208">String</span><span class="sxs-lookup"><span data-stu-id="a46ad-208">String</span></span>|<span data-ttu-id="a46ad-209">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="a46ad-209">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a46ad-210">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-210">Requirements</span></span>

|<span data-ttu-id="a46ad-211">要求</span><span class="sxs-lookup"><span data-stu-id="a46ad-211">Requirement</span></span>| <span data-ttu-id="a46ad-212">值</span><span class="sxs-lookup"><span data-stu-id="a46ad-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="a46ad-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a46ad-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a46ad-214">1.0</span><span class="sxs-lookup"><span data-stu-id="a46ad-214">1.0</span></span>|
|[<span data-ttu-id="a46ad-215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a46ad-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a46ad-216">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a46ad-216">Compose or read</span></span>|