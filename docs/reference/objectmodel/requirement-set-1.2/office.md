 

# <a name="office"></a><span data-ttu-id="eb4d1-101">Office</span><span class="sxs-lookup"><span data-stu-id="eb4d1-101">Office</span></span>

<span data-ttu-id="eb4d1-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb4d1-104">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-104">Requirements</span></span>

|<span data-ttu-id="eb4d1-105">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-105">Requirement</span></span>| <span data-ttu-id="eb4d1-106">值</span><span class="sxs-lookup"><span data-stu-id="eb4d1-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4d1-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb4d1-108">1.0</span><span class="sxs-lookup"><span data-stu-id="eb4d1-108">1.0</span></span>|
|[<span data-ttu-id="eb4d1-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="eb4d1-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eb4d1-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="eb4d1-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="eb4d1-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="eb4d1-111">Namespaces</span></span>

<span data-ttu-id="eb4d1-112">[context](office.context.md)：提供 Office 外接程序 API 的上下文命名空间中的共享接口以便在 Outlook 外接程序 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="eb4d1-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="eb4d1-114">成员</span><span class="sxs-lookup"><span data-stu-id="eb4d1-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="eb4d1-115">AsyncResultStatus： 字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="eb4d1-116">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4d1-117">类型:</span><span class="sxs-lookup"><span data-stu-id="eb4d1-117">Type:</span></span>

*   <span data-ttu-id="eb4d1-118">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eb4d1-119">属性:</span><span class="sxs-lookup"><span data-stu-id="eb4d1-119">Properties:</span></span>

|<span data-ttu-id="eb4d1-120">名称</span><span class="sxs-lookup"><span data-stu-id="eb4d1-120">Name</span></span>| <span data-ttu-id="eb4d1-121">类型</span><span class="sxs-lookup"><span data-stu-id="eb4d1-121">Type</span></span>| <span data-ttu-id="eb4d1-122">说明</span><span class="sxs-lookup"><span data-stu-id="eb4d1-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="eb4d1-123">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-123">String</span></span>|<span data-ttu-id="eb4d1-124">调用成功。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="eb4d1-125">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-125">String</span></span>|<span data-ttu-id="eb4d1-126">调用失败。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb4d1-127">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-127">Requirements</span></span>

|<span data-ttu-id="eb4d1-128">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-128">Requirement</span></span>| <span data-ttu-id="eb4d1-129">值</span><span class="sxs-lookup"><span data-stu-id="eb4d1-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4d1-130">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb4d1-131">1.0</span><span class="sxs-lookup"><span data-stu-id="eb4d1-131">1.0</span></span>|
|[<span data-ttu-id="eb4d1-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="eb4d1-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eb4d1-133">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="eb4d1-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="eb4d1-134">CoercionType： 字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-134">CoercionType :String</span></span>

<span data-ttu-id="eb4d1-135">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4d1-136">类型:</span><span class="sxs-lookup"><span data-stu-id="eb4d1-136">Type:</span></span>

*   <span data-ttu-id="eb4d1-137">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eb4d1-138">属性:</span><span class="sxs-lookup"><span data-stu-id="eb4d1-138">Properties:</span></span>

|<span data-ttu-id="eb4d1-139">名称</span><span class="sxs-lookup"><span data-stu-id="eb4d1-139">Name</span></span>| <span data-ttu-id="eb4d1-140">类型</span><span class="sxs-lookup"><span data-stu-id="eb4d1-140">Type</span></span>| <span data-ttu-id="eb4d1-141">说明</span><span class="sxs-lookup"><span data-stu-id="eb4d1-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="eb4d1-142">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-142">String</span></span>|<span data-ttu-id="eb4d1-143">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="eb4d1-144">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-144">String</span></span>|<span data-ttu-id="eb4d1-145">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb4d1-146">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-146">Requirements</span></span>

|<span data-ttu-id="eb4d1-147">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-147">Requirement</span></span>| <span data-ttu-id="eb4d1-148">值</span><span class="sxs-lookup"><span data-stu-id="eb4d1-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4d1-149">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb4d1-150">1.0</span><span class="sxs-lookup"><span data-stu-id="eb4d1-150">1.0</span></span>|
|[<span data-ttu-id="eb4d1-151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="eb4d1-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eb4d1-152">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="eb4d1-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="eb4d1-153">SourceProperty： 字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-153">SourceProperty :String</span></span>

<span data-ttu-id="eb4d1-154">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4d1-155">类型:</span><span class="sxs-lookup"><span data-stu-id="eb4d1-155">Type:</span></span>

*   <span data-ttu-id="eb4d1-156">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eb4d1-157">属性:</span><span class="sxs-lookup"><span data-stu-id="eb4d1-157">Properties:</span></span>

|<span data-ttu-id="eb4d1-158">名称</span><span class="sxs-lookup"><span data-stu-id="eb4d1-158">Name</span></span>| <span data-ttu-id="eb4d1-159">类型</span><span class="sxs-lookup"><span data-stu-id="eb4d1-159">Type</span></span>| <span data-ttu-id="eb4d1-160">说明</span><span class="sxs-lookup"><span data-stu-id="eb4d1-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="eb4d1-161">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-161">String</span></span>|<span data-ttu-id="eb4d1-162">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="eb4d1-163">字符串</span><span class="sxs-lookup"><span data-stu-id="eb4d1-163">String</span></span>|<span data-ttu-id="eb4d1-164">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="eb4d1-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb4d1-165">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-165">Requirements</span></span>

|<span data-ttu-id="eb4d1-166">要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-166">Requirement</span></span>| <span data-ttu-id="eb4d1-167">值</span><span class="sxs-lookup"><span data-stu-id="eb4d1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4d1-168">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="eb4d1-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb4d1-169">1.0</span><span class="sxs-lookup"><span data-stu-id="eb4d1-169">1.0</span></span>|
|[<span data-ttu-id="eb4d1-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="eb4d1-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eb4d1-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="eb4d1-171">Compose or read</span></span>|