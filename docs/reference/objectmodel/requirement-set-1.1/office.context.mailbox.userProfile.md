
# <a name="userprofile"></a><span data-ttu-id="d63d7-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="d63d7-101">userProfile</span></span>

### <span data-ttu-id="d63d7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md)。userProfile</span><span class="sxs-lookup"><span data-stu-id="d63d7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d63d7-104">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-104">Requirements</span></span>

|<span data-ttu-id="d63d7-105">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-105">Requirement</span></span>| <span data-ttu-id="d63d7-106">值</span><span class="sxs-lookup"><span data-stu-id="d63d7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d63d7-107">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="d63d7-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d63d7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d63d7-108">1.0</span></span>|
|[<span data-ttu-id="d63d7-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d63d7-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d63d7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d63d7-110">ReadItem</span></span>|
|[<span data-ttu-id="d63d7-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d63d7-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d63d7-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d63d7-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="d63d7-113">成员</span><span class="sxs-lookup"><span data-stu-id="d63d7-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="d63d7-114">displayName :字符串</span><span class="sxs-lookup"><span data-stu-id="d63d7-114">displayName :String</span></span>

<span data-ttu-id="d63d7-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="d63d7-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d63d7-116">类型：</span><span class="sxs-lookup"><span data-stu-id="d63d7-116">Type:</span></span>

*   <span data-ttu-id="d63d7-117">字符串</span><span class="sxs-lookup"><span data-stu-id="d63d7-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d63d7-118">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-118">Requirements</span></span>

|<span data-ttu-id="d63d7-119">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-119">Requirement</span></span>| <span data-ttu-id="d63d7-120">值</span><span class="sxs-lookup"><span data-stu-id="d63d7-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="d63d7-121">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="d63d7-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d63d7-122">1.0</span><span class="sxs-lookup"><span data-stu-id="d63d7-122">1.0</span></span>|
|[<span data-ttu-id="d63d7-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d63d7-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d63d7-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d63d7-124">ReadItem</span></span>|
|[<span data-ttu-id="d63d7-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d63d7-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d63d7-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d63d7-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d63d7-127">示例</span><span class="sxs-lookup"><span data-stu-id="d63d7-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d63d7-128">emailAddress :字符串</span><span class="sxs-lookup"><span data-stu-id="d63d7-128">emailAddress :String</span></span>

<span data-ttu-id="d63d7-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d63d7-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d63d7-130">类型：</span><span class="sxs-lookup"><span data-stu-id="d63d7-130">Type:</span></span>

*   <span data-ttu-id="d63d7-131">字符串</span><span class="sxs-lookup"><span data-stu-id="d63d7-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d63d7-132">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-132">Requirements</span></span>

|<span data-ttu-id="d63d7-133">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-133">Requirement</span></span>| <span data-ttu-id="d63d7-134">值</span><span class="sxs-lookup"><span data-stu-id="d63d7-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="d63d7-135">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="d63d7-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d63d7-136">1.0</span><span class="sxs-lookup"><span data-stu-id="d63d7-136">1.0</span></span>|
|[<span data-ttu-id="d63d7-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d63d7-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d63d7-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d63d7-138">ReadItem</span></span>|
|[<span data-ttu-id="d63d7-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d63d7-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d63d7-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d63d7-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d63d7-141">示例</span><span class="sxs-lookup"><span data-stu-id="d63d7-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d63d7-142">timeZone :字符串</span><span class="sxs-lookup"><span data-stu-id="d63d7-142">timeZone :String</span></span>

<span data-ttu-id="d63d7-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="d63d7-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d63d7-144">类型：</span><span class="sxs-lookup"><span data-stu-id="d63d7-144">Type:</span></span>

*   <span data-ttu-id="d63d7-145">字符串</span><span class="sxs-lookup"><span data-stu-id="d63d7-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d63d7-146">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-146">Requirements</span></span>

|<span data-ttu-id="d63d7-147">要求</span><span class="sxs-lookup"><span data-stu-id="d63d7-147">Requirement</span></span>| <span data-ttu-id="d63d7-148">值</span><span class="sxs-lookup"><span data-stu-id="d63d7-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="d63d7-149">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="d63d7-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d63d7-150">1.0</span><span class="sxs-lookup"><span data-stu-id="d63d7-150">1.0</span></span>|
|[<span data-ttu-id="d63d7-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d63d7-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d63d7-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d63d7-152">ReadItem</span></span>|
|[<span data-ttu-id="d63d7-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d63d7-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d63d7-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d63d7-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d63d7-155">示例</span><span class="sxs-lookup"><span data-stu-id="d63d7-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```