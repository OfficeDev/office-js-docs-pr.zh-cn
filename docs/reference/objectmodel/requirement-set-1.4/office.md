 

# <a name="office"></a><span data-ttu-id="93608-101">Office</span><span class="sxs-lookup"><span data-stu-id="93608-101">Office</span></span>

<span data-ttu-id="93608-p101">Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="93608-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="93608-104">要求</span><span class="sxs-lookup"><span data-stu-id="93608-104">Requirements</span></span>

|<span data-ttu-id="93608-105">要求</span><span class="sxs-lookup"><span data-stu-id="93608-105">Requirement</span></span>| <span data-ttu-id="93608-106">值</span><span class="sxs-lookup"><span data-stu-id="93608-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="93608-107">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="93608-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93608-108">1.0</span><span class="sxs-lookup"><span data-stu-id="93608-108">1.0</span></span>|
|[<span data-ttu-id="93608-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="93608-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93608-110">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="93608-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="93608-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="93608-111">Namespaces</span></span>

<span data-ttu-id="93608-112">[上下文](Office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="93608-112">[context](Office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="93608-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="93608-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="93608-114">成员</span><span class="sxs-lookup"><span data-stu-id="93608-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="93608-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="93608-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="93608-116">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="93608-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="93608-117">类型：</span><span class="sxs-lookup"><span data-stu-id="93608-117">Type:</span></span>

*   <span data-ttu-id="93608-118">String</span><span class="sxs-lookup"><span data-stu-id="93608-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93608-119">属性：</span><span class="sxs-lookup"><span data-stu-id="93608-119">Properties:</span></span>

|<span data-ttu-id="93608-120">名称</span><span class="sxs-lookup"><span data-stu-id="93608-120">Name</span></span>| <span data-ttu-id="93608-121">类型</span><span class="sxs-lookup"><span data-stu-id="93608-121">Type</span></span>| <span data-ttu-id="93608-122">说明</span><span class="sxs-lookup"><span data-stu-id="93608-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="93608-123">String</span><span class="sxs-lookup"><span data-stu-id="93608-123">String</span></span>|<span data-ttu-id="93608-124">调用成功。</span><span class="sxs-lookup"><span data-stu-id="93608-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="93608-125">String</span><span class="sxs-lookup"><span data-stu-id="93608-125">String</span></span>|<span data-ttu-id="93608-126">调用失败。</span><span class="sxs-lookup"><span data-stu-id="93608-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93608-127">要求</span><span class="sxs-lookup"><span data-stu-id="93608-127">Requirements</span></span>

|<span data-ttu-id="93608-128">要求</span><span class="sxs-lookup"><span data-stu-id="93608-128">Requirement</span></span>| <span data-ttu-id="93608-129">值</span><span class="sxs-lookup"><span data-stu-id="93608-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="93608-130">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="93608-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93608-131">1.0</span><span class="sxs-lookup"><span data-stu-id="93608-131">1.0</span></span>|
|[<span data-ttu-id="93608-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="93608-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93608-133">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="93608-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="93608-134">CoercionType :字符串</span><span class="sxs-lookup"><span data-stu-id="93608-134">CoercionType :String</span></span>

<span data-ttu-id="93608-135">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="93608-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93608-136">类型：</span><span class="sxs-lookup"><span data-stu-id="93608-136">Type:</span></span>

*   <span data-ttu-id="93608-137">String</span><span class="sxs-lookup"><span data-stu-id="93608-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93608-138">属性：</span><span class="sxs-lookup"><span data-stu-id="93608-138">Properties:</span></span>

|<span data-ttu-id="93608-139">名称</span><span class="sxs-lookup"><span data-stu-id="93608-139">Name</span></span>| <span data-ttu-id="93608-140">类型</span><span class="sxs-lookup"><span data-stu-id="93608-140">Type</span></span>| <span data-ttu-id="93608-141">说明</span><span class="sxs-lookup"><span data-stu-id="93608-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="93608-142">String</span><span class="sxs-lookup"><span data-stu-id="93608-142">String</span></span>|<span data-ttu-id="93608-143">要求以 HTML 格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="93608-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="93608-144">String</span><span class="sxs-lookup"><span data-stu-id="93608-144">String</span></span>|<span data-ttu-id="93608-145">要求以文本格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="93608-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93608-146">要求</span><span class="sxs-lookup"><span data-stu-id="93608-146">Requirements</span></span>

|<span data-ttu-id="93608-147">要求</span><span class="sxs-lookup"><span data-stu-id="93608-147">Requirement</span></span>| <span data-ttu-id="93608-148">值</span><span class="sxs-lookup"><span data-stu-id="93608-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="93608-149">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="93608-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93608-150">1.0</span><span class="sxs-lookup"><span data-stu-id="93608-150">1.0</span></span>|
|[<span data-ttu-id="93608-151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="93608-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93608-152">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="93608-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="93608-153">SourceProperty :字符串</span><span class="sxs-lookup"><span data-stu-id="93608-153">SourceProperty :String</span></span>

<span data-ttu-id="93608-154">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="93608-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93608-155">类型：</span><span class="sxs-lookup"><span data-stu-id="93608-155">Type:</span></span>

*   <span data-ttu-id="93608-156">字符串</span><span class="sxs-lookup"><span data-stu-id="93608-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93608-157">属性：</span><span class="sxs-lookup"><span data-stu-id="93608-157">Properties:</span></span>

|<span data-ttu-id="93608-158">名称</span><span class="sxs-lookup"><span data-stu-id="93608-158">Name</span></span>| <span data-ttu-id="93608-159">类型</span><span class="sxs-lookup"><span data-stu-id="93608-159">Type</span></span>| <span data-ttu-id="93608-160">说明</span><span class="sxs-lookup"><span data-stu-id="93608-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="93608-161">String</span><span class="sxs-lookup"><span data-stu-id="93608-161">String</span></span>|<span data-ttu-id="93608-162">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="93608-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="93608-163">String</span><span class="sxs-lookup"><span data-stu-id="93608-163">String</span></span>|<span data-ttu-id="93608-164">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="93608-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93608-165">要求</span><span class="sxs-lookup"><span data-stu-id="93608-165">Requirements</span></span>

|<span data-ttu-id="93608-166">要求</span><span class="sxs-lookup"><span data-stu-id="93608-166">Requirement</span></span>| <span data-ttu-id="93608-167">值</span><span class="sxs-lookup"><span data-stu-id="93608-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="93608-168">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="93608-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93608-169">1.0</span><span class="sxs-lookup"><span data-stu-id="93608-169">1.0</span></span>|
|[<span data-ttu-id="93608-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="93608-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93608-171">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="93608-171">Compose or read</span></span>|