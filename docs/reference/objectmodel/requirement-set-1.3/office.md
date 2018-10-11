 

# <a name="office"></a><span data-ttu-id="13208-101">Office</span><span class="sxs-lookup"><span data-stu-id="13208-101">Office</span></span>

<span data-ttu-id="13208-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="13208-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="13208-104">要求</span><span class="sxs-lookup"><span data-stu-id="13208-104">Requirements</span></span>

|<span data-ttu-id="13208-105">要求</span><span class="sxs-lookup"><span data-stu-id="13208-105">Requirement</span></span>| <span data-ttu-id="13208-106">值</span><span class="sxs-lookup"><span data-stu-id="13208-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="13208-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="13208-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13208-108">1.0</span><span class="sxs-lookup"><span data-stu-id="13208-108">1.0</span></span>|
|[<span data-ttu-id="13208-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13208-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13208-110">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="13208-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="13208-111">命名空间</span><span class="sxs-lookup"><span data-stu-id="13208-111">Namespaces</span></span>

<span data-ttu-id="13208-112">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="13208-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="13208-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="13208-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="13208-114">成员</span><span class="sxs-lookup"><span data-stu-id="13208-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="13208-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="13208-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="13208-116">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="13208-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="13208-117">类型：</span><span class="sxs-lookup"><span data-stu-id="13208-117">Type:</span></span>

*   <span data-ttu-id="13208-118">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="13208-119">属性:</span><span class="sxs-lookup"><span data-stu-id="13208-119">Properties:</span></span>

|<span data-ttu-id="13208-120">名称</span><span class="sxs-lookup"><span data-stu-id="13208-120">Name</span></span>| <span data-ttu-id="13208-121">类型</span><span class="sxs-lookup"><span data-stu-id="13208-121">Type</span></span>| <span data-ttu-id="13208-122">说明</span><span class="sxs-lookup"><span data-stu-id="13208-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="13208-123">String</span><span class="sxs-lookup"><span data-stu-id="13208-123">String</span></span>|<span data-ttu-id="13208-124">调用成功。</span><span class="sxs-lookup"><span data-stu-id="13208-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="13208-125">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-125">String</span></span>|<span data-ttu-id="13208-126">调用失败。</span><span class="sxs-lookup"><span data-stu-id="13208-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13208-127">要求</span><span class="sxs-lookup"><span data-stu-id="13208-127">Requirements</span></span>

|<span data-ttu-id="13208-128">要求</span><span class="sxs-lookup"><span data-stu-id="13208-128">Requirement</span></span>| <span data-ttu-id="13208-129">值</span><span class="sxs-lookup"><span data-stu-id="13208-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="13208-130">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="13208-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13208-131">1.0</span><span class="sxs-lookup"><span data-stu-id="13208-131">1.0</span></span>|
|[<span data-ttu-id="13208-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13208-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13208-133">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="13208-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="13208-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="13208-134">CoercionType :String</span></span>

<span data-ttu-id="13208-135">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="13208-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="13208-136">类型：</span><span class="sxs-lookup"><span data-stu-id="13208-136">Type:</span></span>

*   <span data-ttu-id="13208-137">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="13208-138">属性:</span><span class="sxs-lookup"><span data-stu-id="13208-138">Properties:</span></span>

|<span data-ttu-id="13208-139">名称</span><span class="sxs-lookup"><span data-stu-id="13208-139">Name</span></span>| <span data-ttu-id="13208-140">类型</span><span class="sxs-lookup"><span data-stu-id="13208-140">Type</span></span>| <span data-ttu-id="13208-141">说明</span><span class="sxs-lookup"><span data-stu-id="13208-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="13208-142">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-142">String</span></span>|<span data-ttu-id="13208-143">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="13208-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="13208-144">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-144">String</span></span>|<span data-ttu-id="13208-145">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="13208-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13208-146">要求</span><span class="sxs-lookup"><span data-stu-id="13208-146">Requirements</span></span>

|<span data-ttu-id="13208-147">要求</span><span class="sxs-lookup"><span data-stu-id="13208-147">Requirement</span></span>| <span data-ttu-id="13208-148">值</span><span class="sxs-lookup"><span data-stu-id="13208-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="13208-149">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="13208-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13208-150">1.0</span><span class="sxs-lookup"><span data-stu-id="13208-150">1.0</span></span>|
|[<span data-ttu-id="13208-151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13208-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13208-152">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="13208-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="13208-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="13208-153">SourceProperty :String</span></span>

<span data-ttu-id="13208-154">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="13208-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="13208-155">类型：</span><span class="sxs-lookup"><span data-stu-id="13208-155">Type:</span></span>

*   <span data-ttu-id="13208-156">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="13208-157">属性:</span><span class="sxs-lookup"><span data-stu-id="13208-157">Properties:</span></span>

|<span data-ttu-id="13208-158">名称</span><span class="sxs-lookup"><span data-stu-id="13208-158">Name</span></span>| <span data-ttu-id="13208-159">类型</span><span class="sxs-lookup"><span data-stu-id="13208-159">Type</span></span>| <span data-ttu-id="13208-160">说明</span><span class="sxs-lookup"><span data-stu-id="13208-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="13208-161">String</span><span class="sxs-lookup"><span data-stu-id="13208-161">String</span></span>|<span data-ttu-id="13208-162">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="13208-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="13208-163">字符串</span><span class="sxs-lookup"><span data-stu-id="13208-163">String</span></span>|<span data-ttu-id="13208-164">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="13208-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13208-165">要求</span><span class="sxs-lookup"><span data-stu-id="13208-165">Requirements</span></span>

|<span data-ttu-id="13208-166">要求</span><span class="sxs-lookup"><span data-stu-id="13208-166">Requirement</span></span>| <span data-ttu-id="13208-167">值</span><span class="sxs-lookup"><span data-stu-id="13208-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="13208-168">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="13208-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13208-169">1.0</span><span class="sxs-lookup"><span data-stu-id="13208-169">1.0</span></span>|
|[<span data-ttu-id="13208-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13208-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13208-171">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="13208-171">Compose or read</span></span>|