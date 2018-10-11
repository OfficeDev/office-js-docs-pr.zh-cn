 

# <a name="office"></a><span data-ttu-id="51e13-101">Office</span><span class="sxs-lookup"><span data-stu-id="51e13-101">Office</span></span>

<span data-ttu-id="51e13-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="51e13-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="51e13-104">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-104">Requirements</span></span>

|<span data-ttu-id="51e13-105">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-105">Requirement</span></span>| <span data-ttu-id="51e13-106">值</span><span class="sxs-lookup"><span data-stu-id="51e13-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="51e13-107">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="51e13-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51e13-108">1.0</span><span class="sxs-lookup"><span data-stu-id="51e13-108">1.0</span></span>|
|[<span data-ttu-id="51e13-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="51e13-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="51e13-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="51e13-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="51e13-111">命名空间</span><span class="sxs-lookup"><span data-stu-id="51e13-111">Namespaces</span></span>

<span data-ttu-id="51e13-112">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="51e13-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="51e13-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="51e13-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="51e13-114">成员</span><span class="sxs-lookup"><span data-stu-id="51e13-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="51e13-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="51e13-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="51e13-116">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="51e13-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="51e13-117">类型：</span><span class="sxs-lookup"><span data-stu-id="51e13-117">Type:</span></span>

*   <span data-ttu-id="51e13-118">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="51e13-119">属性：</span><span class="sxs-lookup"><span data-stu-id="51e13-119">Properties:</span></span>

|<span data-ttu-id="51e13-120">名称</span><span class="sxs-lookup"><span data-stu-id="51e13-120">Name</span></span>| <span data-ttu-id="51e13-121">类型</span><span class="sxs-lookup"><span data-stu-id="51e13-121">Type</span></span>| <span data-ttu-id="51e13-122">说明</span><span class="sxs-lookup"><span data-stu-id="51e13-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="51e13-123">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-123">String</span></span>|<span data-ttu-id="51e13-124">调用成功。</span><span class="sxs-lookup"><span data-stu-id="51e13-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="51e13-125">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-125">String</span></span>|<span data-ttu-id="51e13-126">调用失败。</span><span class="sxs-lookup"><span data-stu-id="51e13-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="51e13-127">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-127">Requirements</span></span>

|<span data-ttu-id="51e13-128">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-128">Requirement</span></span>| <span data-ttu-id="51e13-129">值</span><span class="sxs-lookup"><span data-stu-id="51e13-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="51e13-130">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="51e13-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51e13-131">1.0</span><span class="sxs-lookup"><span data-stu-id="51e13-131">1.0</span></span>|
|[<span data-ttu-id="51e13-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="51e13-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="51e13-133">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="51e13-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="51e13-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="51e13-134">CoercionType :String</span></span>

<span data-ttu-id="51e13-135">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="51e13-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="51e13-136">类型：</span><span class="sxs-lookup"><span data-stu-id="51e13-136">Type:</span></span>

*   <span data-ttu-id="51e13-137">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="51e13-138">属性：</span><span class="sxs-lookup"><span data-stu-id="51e13-138">Properties:</span></span>

|<span data-ttu-id="51e13-139">名称</span><span class="sxs-lookup"><span data-stu-id="51e13-139">Name</span></span>| <span data-ttu-id="51e13-140">类型</span><span class="sxs-lookup"><span data-stu-id="51e13-140">Type</span></span>| <span data-ttu-id="51e13-141">说明</span><span class="sxs-lookup"><span data-stu-id="51e13-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="51e13-142">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-142">String</span></span>|<span data-ttu-id="51e13-143">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="51e13-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="51e13-144">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-144">String</span></span>|<span data-ttu-id="51e13-145">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="51e13-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="51e13-146">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-146">Requirements</span></span>

|<span data-ttu-id="51e13-147">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-147">Requirement</span></span>| <span data-ttu-id="51e13-148">值</span><span class="sxs-lookup"><span data-stu-id="51e13-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="51e13-149">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="51e13-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51e13-150">1.0</span><span class="sxs-lookup"><span data-stu-id="51e13-150">1.0</span></span>|
|[<span data-ttu-id="51e13-151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="51e13-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="51e13-152">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="51e13-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="51e13-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="51e13-153">SourceProperty :String</span></span>

<span data-ttu-id="51e13-154">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="51e13-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="51e13-155">类型：</span><span class="sxs-lookup"><span data-stu-id="51e13-155">Type:</span></span>

*   <span data-ttu-id="51e13-156">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="51e13-157">属性：</span><span class="sxs-lookup"><span data-stu-id="51e13-157">Properties:</span></span>

|<span data-ttu-id="51e13-158">名称</span><span class="sxs-lookup"><span data-stu-id="51e13-158">Name</span></span>| <span data-ttu-id="51e13-159">类型</span><span class="sxs-lookup"><span data-stu-id="51e13-159">Type</span></span>| <span data-ttu-id="51e13-160">说明</span><span class="sxs-lookup"><span data-stu-id="51e13-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="51e13-161">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-161">String</span></span>|<span data-ttu-id="51e13-162">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="51e13-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="51e13-163">字符串</span><span class="sxs-lookup"><span data-stu-id="51e13-163">String</span></span>|<span data-ttu-id="51e13-164">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="51e13-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="51e13-165">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-165">Requirements</span></span>

|<span data-ttu-id="51e13-166">要求</span><span class="sxs-lookup"><span data-stu-id="51e13-166">Requirement</span></span>| <span data-ttu-id="51e13-167">值</span><span class="sxs-lookup"><span data-stu-id="51e13-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="51e13-168">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="51e13-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51e13-169">1.0</span><span class="sxs-lookup"><span data-stu-id="51e13-169">1.0</span></span>|
|[<span data-ttu-id="51e13-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="51e13-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="51e13-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="51e13-171">Compose or read</span></span>|