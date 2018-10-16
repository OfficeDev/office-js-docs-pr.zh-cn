# <a name="office"></a><span data-ttu-id="f529d-101">Office</span><span class="sxs-lookup"><span data-stu-id="f529d-101">Office</span></span>

<span data-ttu-id="f529d-p101">Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="f529d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f529d-104">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-104">Requirements</span></span>

|<span data-ttu-id="f529d-105">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-105">Requirement</span></span>| <span data-ttu-id="f529d-106">值</span><span class="sxs-lookup"><span data-stu-id="f529d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f529d-107">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f529d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f529d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f529d-108">1.0</span></span>|
|[<span data-ttu-id="f529d-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f529d-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f529d-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f529d-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f529d-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="f529d-111">Members and methods</span></span>

| <span data-ttu-id="f529d-112">成员</span><span class="sxs-lookup"><span data-stu-id="f529d-112">Member</span></span> | <span data-ttu-id="f529d-113">类型</span><span class="sxs-lookup"><span data-stu-id="f529d-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f529d-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f529d-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f529d-115">成员</span><span class="sxs-lookup"><span data-stu-id="f529d-115">Member</span></span> |
| [<span data-ttu-id="f529d-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f529d-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f529d-117">成员</span><span class="sxs-lookup"><span data-stu-id="f529d-117">Member</span></span> |
| [<span data-ttu-id="f529d-118">EventType</span><span class="sxs-lookup"><span data-stu-id="f529d-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f529d-119">成员</span><span class="sxs-lookup"><span data-stu-id="f529d-119">Member</span></span> |
| [<span data-ttu-id="f529d-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f529d-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f529d-121">成员</span><span class="sxs-lookup"><span data-stu-id="f529d-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f529d-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f529d-122">Namespaces</span></span>

<span data-ttu-id="f529d-123">[上下文](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="f529d-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f529d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="f529d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f529d-125">成员</span><span class="sxs-lookup"><span data-stu-id="f529d-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f529d-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f529d-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="f529d-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="f529d-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f529d-128">类型：</span><span class="sxs-lookup"><span data-stu-id="f529d-128">Type:</span></span>

*   <span data-ttu-id="f529d-129">String</span><span class="sxs-lookup"><span data-stu-id="f529d-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f529d-130">属性：</span><span class="sxs-lookup"><span data-stu-id="f529d-130">Properties:</span></span>

|<span data-ttu-id="f529d-131">名称</span><span class="sxs-lookup"><span data-stu-id="f529d-131">Name</span></span>| <span data-ttu-id="f529d-132">类型</span><span class="sxs-lookup"><span data-stu-id="f529d-132">Type</span></span>| <span data-ttu-id="f529d-133">说明</span><span class="sxs-lookup"><span data-stu-id="f529d-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f529d-134">String</span><span class="sxs-lookup"><span data-stu-id="f529d-134">String</span></span>|<span data-ttu-id="f529d-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="f529d-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f529d-136">String</span><span class="sxs-lookup"><span data-stu-id="f529d-136">String</span></span>|<span data-ttu-id="f529d-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="f529d-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f529d-138">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-138">Requirements</span></span>

|<span data-ttu-id="f529d-139">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-139">Requirement</span></span>| <span data-ttu-id="f529d-140">值</span><span class="sxs-lookup"><span data-stu-id="f529d-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="f529d-141">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f529d-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f529d-142">1.0</span><span class="sxs-lookup"><span data-stu-id="f529d-142">1.0</span></span>|
|[<span data-ttu-id="f529d-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f529d-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f529d-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f529d-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f529d-145">CoercionType :字符串</span><span class="sxs-lookup"><span data-stu-id="f529d-145">CoercionType :String</span></span>

<span data-ttu-id="f529d-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="f529d-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f529d-147">类型：</span><span class="sxs-lookup"><span data-stu-id="f529d-147">Type:</span></span>

*   <span data-ttu-id="f529d-148">String</span><span class="sxs-lookup"><span data-stu-id="f529d-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f529d-149">属性：</span><span class="sxs-lookup"><span data-stu-id="f529d-149">Properties:</span></span>

|<span data-ttu-id="f529d-150">名称</span><span class="sxs-lookup"><span data-stu-id="f529d-150">Name</span></span>| <span data-ttu-id="f529d-151">类型</span><span class="sxs-lookup"><span data-stu-id="f529d-151">Type</span></span>| <span data-ttu-id="f529d-152">说明</span><span class="sxs-lookup"><span data-stu-id="f529d-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f529d-153">String</span><span class="sxs-lookup"><span data-stu-id="f529d-153">String</span></span>|<span data-ttu-id="f529d-154">要求以 HTML 格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="f529d-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f529d-155">String</span><span class="sxs-lookup"><span data-stu-id="f529d-155">String</span></span>|<span data-ttu-id="f529d-156">要求以文本格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="f529d-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f529d-157">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-157">Requirements</span></span>

|<span data-ttu-id="f529d-158">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-158">Requirement</span></span>| <span data-ttu-id="f529d-159">值</span><span class="sxs-lookup"><span data-stu-id="f529d-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="f529d-160">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f529d-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f529d-161">1.0</span><span class="sxs-lookup"><span data-stu-id="f529d-161">1.0</span></span>|
|[<span data-ttu-id="f529d-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f529d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f529d-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f529d-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f529d-164">EventType :字符串</span><span class="sxs-lookup"><span data-stu-id="f529d-164">EventType :String</span></span>

<span data-ttu-id="f529d-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="f529d-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f529d-166">类型：</span><span class="sxs-lookup"><span data-stu-id="f529d-166">Type:</span></span>

*   <span data-ttu-id="f529d-167">String</span><span class="sxs-lookup"><span data-stu-id="f529d-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f529d-168">属性：</span><span class="sxs-lookup"><span data-stu-id="f529d-168">Properties:</span></span>

| <span data-ttu-id="f529d-169">名称</span><span class="sxs-lookup"><span data-stu-id="f529d-169">Name</span></span> | <span data-ttu-id="f529d-170">类型</span><span class="sxs-lookup"><span data-stu-id="f529d-170">Type</span></span> | <span data-ttu-id="f529d-171">说明</span><span class="sxs-lookup"><span data-stu-id="f529d-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="f529d-172">String</span><span class="sxs-lookup"><span data-stu-id="f529d-172">String</span></span> | <span data-ttu-id="f529d-173">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="f529d-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f529d-174">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-174">Requirements</span></span>

|<span data-ttu-id="f529d-175">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-175">Requirement</span></span>| <span data-ttu-id="f529d-176">值</span><span class="sxs-lookup"><span data-stu-id="f529d-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="f529d-177">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f529d-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f529d-178">1.5</span><span class="sxs-lookup"><span data-stu-id="f529d-178">1.5</span></span> |
|[<span data-ttu-id="f529d-179">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f529d-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f529d-180">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f529d-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f529d-181">SourceProperty :字符串</span><span class="sxs-lookup"><span data-stu-id="f529d-181">SourceProperty :String</span></span>

<span data-ttu-id="f529d-182">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="f529d-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f529d-183">类型：</span><span class="sxs-lookup"><span data-stu-id="f529d-183">Type:</span></span>

*   <span data-ttu-id="f529d-184">String</span><span class="sxs-lookup"><span data-stu-id="f529d-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f529d-185">属性：</span><span class="sxs-lookup"><span data-stu-id="f529d-185">Properties:</span></span>

|<span data-ttu-id="f529d-186">名称</span><span class="sxs-lookup"><span data-stu-id="f529d-186">Name</span></span>| <span data-ttu-id="f529d-187">类型</span><span class="sxs-lookup"><span data-stu-id="f529d-187">Type</span></span>| <span data-ttu-id="f529d-188">说明</span><span class="sxs-lookup"><span data-stu-id="f529d-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f529d-189">String</span><span class="sxs-lookup"><span data-stu-id="f529d-189">String</span></span>|<span data-ttu-id="f529d-190">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="f529d-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f529d-191">String</span><span class="sxs-lookup"><span data-stu-id="f529d-191">String</span></span>|<span data-ttu-id="f529d-192">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="f529d-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f529d-193">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-193">Requirements</span></span>

|<span data-ttu-id="f529d-194">要求</span><span class="sxs-lookup"><span data-stu-id="f529d-194">Requirement</span></span>| <span data-ttu-id="f529d-195">值</span><span class="sxs-lookup"><span data-stu-id="f529d-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="f529d-196">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f529d-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f529d-197">1.0</span><span class="sxs-lookup"><span data-stu-id="f529d-197">1.0</span></span>|
|[<span data-ttu-id="f529d-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f529d-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f529d-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f529d-199">Compose or read</span></span>|