# <a name="office"></a><span data-ttu-id="0ca8c-101">Office</span><span class="sxs-lookup"><span data-stu-id="0ca8c-101">Office</span></span>

<span data-ttu-id="0ca8c-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ca8c-104">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-104">Requirements</span></span>

|<span data-ttu-id="0ca8c-105">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-105">Requirement</span></span>| <span data-ttu-id="0ca8c-106">值</span><span class="sxs-lookup"><span data-stu-id="0ca8c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ca8c-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ca8c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0ca8c-108">1.0</span></span>|
|[<span data-ttu-id="0ca8c-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ca8c-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0ca8c-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ca8c-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0ca8c-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="0ca8c-111">Members and methods</span></span>

| <span data-ttu-id="0ca8c-112">成员</span><span class="sxs-lookup"><span data-stu-id="0ca8c-112">Member</span></span> | <span data-ttu-id="0ca8c-113">类型</span><span class="sxs-lookup"><span data-stu-id="0ca8c-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0ca8c-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0ca8c-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0ca8c-115">成员</span><span class="sxs-lookup"><span data-stu-id="0ca8c-115">Member</span></span> |
| [<span data-ttu-id="0ca8c-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0ca8c-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0ca8c-117">成员</span><span class="sxs-lookup"><span data-stu-id="0ca8c-117">Member</span></span> |
| [<span data-ttu-id="0ca8c-118">EventType</span><span class="sxs-lookup"><span data-stu-id="0ca8c-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0ca8c-119">成员</span><span class="sxs-lookup"><span data-stu-id="0ca8c-119">Member</span></span> |
| [<span data-ttu-id="0ca8c-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0ca8c-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0ca8c-121">成员</span><span class="sxs-lookup"><span data-stu-id="0ca8c-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0ca8c-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="0ca8c-122">Namespaces</span></span>

<span data-ttu-id="0ca8c-123">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0ca8c-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="0ca8c-125">成员</span><span class="sxs-lookup"><span data-stu-id="0ca8c-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="0ca8c-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="0ca8c-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="0ca8c-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0ca8c-128">类型：</span><span class="sxs-lookup"><span data-stu-id="0ca8c-128">Type:</span></span>

*   <span data-ttu-id="0ca8c-129">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ca8c-130">属性:</span><span class="sxs-lookup"><span data-stu-id="0ca8c-130">Properties:</span></span>

|<span data-ttu-id="0ca8c-131">名称</span><span class="sxs-lookup"><span data-stu-id="0ca8c-131">Name</span></span>| <span data-ttu-id="0ca8c-132">类型</span><span class="sxs-lookup"><span data-stu-id="0ca8c-132">Type</span></span>| <span data-ttu-id="0ca8c-133">说明</span><span class="sxs-lookup"><span data-stu-id="0ca8c-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0ca8c-134">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-134">String</span></span>|<span data-ttu-id="0ca8c-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0ca8c-136">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-136">String</span></span>|<span data-ttu-id="0ca8c-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ca8c-138">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-138">Requirements</span></span>

|<span data-ttu-id="0ca8c-139">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-139">Requirement</span></span>| <span data-ttu-id="0ca8c-140">值</span><span class="sxs-lookup"><span data-stu-id="0ca8c-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ca8c-141">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ca8c-142">1.0</span><span class="sxs-lookup"><span data-stu-id="0ca8c-142">1.0</span></span>|
|[<span data-ttu-id="0ca8c-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ca8c-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0ca8c-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ca8c-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="0ca8c-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="0ca8c-145">CoercionType :String</span></span>

<span data-ttu-id="0ca8c-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0ca8c-147">类型：</span><span class="sxs-lookup"><span data-stu-id="0ca8c-147">Type:</span></span>

*   <span data-ttu-id="0ca8c-148">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ca8c-149">属性:</span><span class="sxs-lookup"><span data-stu-id="0ca8c-149">Properties:</span></span>

|<span data-ttu-id="0ca8c-150">名称</span><span class="sxs-lookup"><span data-stu-id="0ca8c-150">Name</span></span>| <span data-ttu-id="0ca8c-151">类型</span><span class="sxs-lookup"><span data-stu-id="0ca8c-151">Type</span></span>| <span data-ttu-id="0ca8c-152">说明</span><span class="sxs-lookup"><span data-stu-id="0ca8c-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0ca8c-153">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-153">String</span></span>|<span data-ttu-id="0ca8c-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0ca8c-155">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-155">String</span></span>|<span data-ttu-id="0ca8c-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ca8c-157">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-157">Requirements</span></span>

|<span data-ttu-id="0ca8c-158">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-158">Requirement</span></span>| <span data-ttu-id="0ca8c-159">值</span><span class="sxs-lookup"><span data-stu-id="0ca8c-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ca8c-160">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ca8c-161">1.0</span><span class="sxs-lookup"><span data-stu-id="0ca8c-161">1.0</span></span>|
|[<span data-ttu-id="0ca8c-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ca8c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0ca8c-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ca8c-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="0ca8c-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="0ca8c-164">EventType :String</span></span>

<span data-ttu-id="0ca8c-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0ca8c-166">类型：</span><span class="sxs-lookup"><span data-stu-id="0ca8c-166">Type:</span></span>

*   <span data-ttu-id="0ca8c-167">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ca8c-168">属性:</span><span class="sxs-lookup"><span data-stu-id="0ca8c-168">Properties:</span></span>

| <span data-ttu-id="0ca8c-169">名称</span><span class="sxs-lookup"><span data-stu-id="0ca8c-169">Name</span></span> | <span data-ttu-id="0ca8c-170">类型</span><span class="sxs-lookup"><span data-stu-id="0ca8c-170">Type</span></span> | <span data-ttu-id="0ca8c-171">说明</span><span class="sxs-lookup"><span data-stu-id="0ca8c-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="0ca8c-172">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-172">String</span></span> | <span data-ttu-id="0ca8c-173">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0ca8c-174">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-174">Requirements</span></span>

|<span data-ttu-id="0ca8c-175">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-175">Requirement</span></span>| <span data-ttu-id="0ca8c-176">值</span><span class="sxs-lookup"><span data-stu-id="0ca8c-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ca8c-177">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ca8c-178">1.5</span><span class="sxs-lookup"><span data-stu-id="0ca8c-178">1.5</span></span> |
|[<span data-ttu-id="0ca8c-179">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ca8c-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0ca8c-180">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ca8c-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="0ca8c-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="0ca8c-181">SourceProperty :String</span></span>

<span data-ttu-id="0ca8c-182">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0ca8c-183">类型：</span><span class="sxs-lookup"><span data-stu-id="0ca8c-183">Type:</span></span>

*   <span data-ttu-id="0ca8c-184">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ca8c-185">属性:</span><span class="sxs-lookup"><span data-stu-id="0ca8c-185">Properties:</span></span>

|<span data-ttu-id="0ca8c-186">名称</span><span class="sxs-lookup"><span data-stu-id="0ca8c-186">Name</span></span>| <span data-ttu-id="0ca8c-187">类型</span><span class="sxs-lookup"><span data-stu-id="0ca8c-187">Type</span></span>| <span data-ttu-id="0ca8c-188">说明</span><span class="sxs-lookup"><span data-stu-id="0ca8c-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0ca8c-189">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-189">String</span></span>|<span data-ttu-id="0ca8c-190">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0ca8c-191">字符串</span><span class="sxs-lookup"><span data-stu-id="0ca8c-191">String</span></span>|<span data-ttu-id="0ca8c-192">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="0ca8c-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ca8c-193">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-193">Requirements</span></span>

|<span data-ttu-id="0ca8c-194">要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-194">Requirement</span></span>| <span data-ttu-id="0ca8c-195">值</span><span class="sxs-lookup"><span data-stu-id="0ca8c-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ca8c-196">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="0ca8c-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ca8c-197">1.0</span><span class="sxs-lookup"><span data-stu-id="0ca8c-197">1.0</span></span>|
|[<span data-ttu-id="0ca8c-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ca8c-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0ca8c-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ca8c-199">Compose or read</span></span>|