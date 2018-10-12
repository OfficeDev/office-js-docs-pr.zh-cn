 

# <a name="office"></a><span data-ttu-id="cdc96-101">Office</span><span class="sxs-lookup"><span data-stu-id="cdc96-101">Office</span></span>

<span data-ttu-id="cdc96-p101">Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="cdc96-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cdc96-104">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-104">Requirements</span></span>

|<span data-ttu-id="cdc96-105">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-105">Requirement</span></span>| <span data-ttu-id="cdc96-106">值</span><span class="sxs-lookup"><span data-stu-id="cdc96-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="cdc96-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cdc96-108">1.0</span><span class="sxs-lookup"><span data-stu-id="cdc96-108">1.0</span></span>|
|[<span data-ttu-id="cdc96-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cdc96-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cdc96-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cdc96-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="cdc96-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="cdc96-111">Namespaces</span></span>

<span data-ttu-id="cdc96-112">[上下文](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="cdc96-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="cdc96-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="cdc96-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="cdc96-114">成员</span><span class="sxs-lookup"><span data-stu-id="cdc96-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="cdc96-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="cdc96-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="cdc96-116">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="cdc96-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cdc96-117">类型：</span><span class="sxs-lookup"><span data-stu-id="cdc96-117">Type:</span></span>

*   <span data-ttu-id="cdc96-118">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cdc96-119">属性：</span><span class="sxs-lookup"><span data-stu-id="cdc96-119">Properties:</span></span>

|<span data-ttu-id="cdc96-120">名称</span><span class="sxs-lookup"><span data-stu-id="cdc96-120">Name</span></span>| <span data-ttu-id="cdc96-121">类型</span><span class="sxs-lookup"><span data-stu-id="cdc96-121">Type</span></span>| <span data-ttu-id="cdc96-122">说明</span><span class="sxs-lookup"><span data-stu-id="cdc96-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cdc96-123">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-123">String</span></span>|<span data-ttu-id="cdc96-124">调用成功。</span><span class="sxs-lookup"><span data-stu-id="cdc96-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cdc96-125">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-125">String</span></span>|<span data-ttu-id="cdc96-126">调用失败。</span><span class="sxs-lookup"><span data-stu-id="cdc96-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cdc96-127">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-127">Requirements</span></span>

|<span data-ttu-id="cdc96-128">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-128">Requirement</span></span>| <span data-ttu-id="cdc96-129">值</span><span class="sxs-lookup"><span data-stu-id="cdc96-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="cdc96-130">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cdc96-131">1.0</span><span class="sxs-lookup"><span data-stu-id="cdc96-131">1.0</span></span>|
|[<span data-ttu-id="cdc96-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cdc96-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cdc96-133">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cdc96-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="cdc96-134">CoercionType :字符串</span><span class="sxs-lookup"><span data-stu-id="cdc96-134">CoercionType :String</span></span>

<span data-ttu-id="cdc96-135">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="cdc96-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cdc96-136">类型：</span><span class="sxs-lookup"><span data-stu-id="cdc96-136">Type:</span></span>

*   <span data-ttu-id="cdc96-137">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cdc96-138">属性：</span><span class="sxs-lookup"><span data-stu-id="cdc96-138">Properties:</span></span>

|<span data-ttu-id="cdc96-139">名称</span><span class="sxs-lookup"><span data-stu-id="cdc96-139">Name</span></span>| <span data-ttu-id="cdc96-140">类型</span><span class="sxs-lookup"><span data-stu-id="cdc96-140">Type</span></span>| <span data-ttu-id="cdc96-141">说明</span><span class="sxs-lookup"><span data-stu-id="cdc96-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cdc96-142">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-142">String</span></span>|<span data-ttu-id="cdc96-143">要求以 HTML 格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="cdc96-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cdc96-144">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-144">String</span></span>|<span data-ttu-id="cdc96-145">要求以文本格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="cdc96-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cdc96-146">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-146">Requirements</span></span>

|<span data-ttu-id="cdc96-147">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-147">Requirement</span></span>| <span data-ttu-id="cdc96-148">值</span><span class="sxs-lookup"><span data-stu-id="cdc96-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="cdc96-149">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cdc96-150">1.0</span><span class="sxs-lookup"><span data-stu-id="cdc96-150">1.0</span></span>|
|[<span data-ttu-id="cdc96-151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cdc96-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cdc96-152">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cdc96-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="cdc96-153">SourceProperty :字符串</span><span class="sxs-lookup"><span data-stu-id="cdc96-153">SourceProperty :String</span></span>

<span data-ttu-id="cdc96-154">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="cdc96-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cdc96-155">类型：</span><span class="sxs-lookup"><span data-stu-id="cdc96-155">Type:</span></span>

*   <span data-ttu-id="cdc96-156">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cdc96-157">属性：</span><span class="sxs-lookup"><span data-stu-id="cdc96-157">Properties:</span></span>

|<span data-ttu-id="cdc96-158">名称</span><span class="sxs-lookup"><span data-stu-id="cdc96-158">Name</span></span>| <span data-ttu-id="cdc96-159">类型</span><span class="sxs-lookup"><span data-stu-id="cdc96-159">Type</span></span>| <span data-ttu-id="cdc96-160">说明</span><span class="sxs-lookup"><span data-stu-id="cdc96-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cdc96-161">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-161">String</span></span>|<span data-ttu-id="cdc96-162">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="cdc96-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cdc96-163">String</span><span class="sxs-lookup"><span data-stu-id="cdc96-163">String</span></span>|<span data-ttu-id="cdc96-164">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="cdc96-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cdc96-165">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-165">Requirements</span></span>

|<span data-ttu-id="cdc96-166">要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-166">Requirement</span></span>| <span data-ttu-id="cdc96-167">值</span><span class="sxs-lookup"><span data-stu-id="cdc96-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cdc96-168">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="cdc96-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cdc96-169">1.0</span><span class="sxs-lookup"><span data-stu-id="cdc96-169">1.0</span></span>|
|[<span data-ttu-id="cdc96-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cdc96-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cdc96-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cdc96-171">Compose or read</span></span>|