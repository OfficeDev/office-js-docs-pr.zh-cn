# <a name="userprofile"></a><span data-ttu-id="a09c5-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="a09c5-101">userProfile</span></span>

### <span data-ttu-id="a09c5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="a09c5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a09c5-104">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-104">Requirements</span></span>

|<span data-ttu-id="a09c5-105">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-105">Requirement</span></span>| <span data-ttu-id="a09c5-106">值</span><span class="sxs-lookup"><span data-stu-id="a09c5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a09c5-107">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a09c5-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a09c5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a09c5-108">1.0</span></span>|
|[<span data-ttu-id="a09c5-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a09c5-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a09c5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a09c5-110">ReadItem</span></span>|
|[<span data-ttu-id="a09c5-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a09c5-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a09c5-112">Compose or read</span><span class="sxs-lookup"><span data-stu-id="a09c5-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a09c5-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="a09c5-113">Members and methods</span></span>

| <span data-ttu-id="a09c5-114">Member</span><span class="sxs-lookup"><span data-stu-id="a09c5-114">Member</span></span> | <span data-ttu-id="a09c5-115">类型</span><span class="sxs-lookup"><span data-stu-id="a09c5-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a09c5-116">displayName</span><span class="sxs-lookup"><span data-stu-id="a09c5-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="a09c5-117">Member</span><span class="sxs-lookup"><span data-stu-id="a09c5-117">Member</span></span> |
| [<span data-ttu-id="a09c5-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a09c5-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a09c5-119">Member</span><span class="sxs-lookup"><span data-stu-id="a09c5-119">Member</span></span> |
| [<span data-ttu-id="a09c5-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="a09c5-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a09c5-121">Member</span><span class="sxs-lookup"><span data-stu-id="a09c5-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a09c5-122">成员</span><span class="sxs-lookup"><span data-stu-id="a09c5-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a09c5-123">displayName :字符串</span><span class="sxs-lookup"><span data-stu-id="a09c5-123">displayName :String</span></span>

<span data-ttu-id="a09c5-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="a09c5-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a09c5-125">类型：</span><span class="sxs-lookup"><span data-stu-id="a09c5-125">Type:</span></span>

*   <span data-ttu-id="a09c5-126">String</span><span class="sxs-lookup"><span data-stu-id="a09c5-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a09c5-127">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-127">Requirements</span></span>

|<span data-ttu-id="a09c5-128">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-128">Requirement</span></span>| <span data-ttu-id="a09c5-129">值</span><span class="sxs-lookup"><span data-stu-id="a09c5-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="a09c5-130">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a09c5-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a09c5-131">1.0</span><span class="sxs-lookup"><span data-stu-id="a09c5-131">1.0</span></span>|
|[<span data-ttu-id="a09c5-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a09c5-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a09c5-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a09c5-133">ReadItem</span></span>|
|[<span data-ttu-id="a09c5-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a09c5-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a09c5-135">Compose or read</span><span class="sxs-lookup"><span data-stu-id="a09c5-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a09c5-136">示例</span><span class="sxs-lookup"><span data-stu-id="a09c5-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a09c5-137">emailAddress :字符串</span><span class="sxs-lookup"><span data-stu-id="a09c5-137">emailAddress :String</span></span>

<span data-ttu-id="a09c5-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a09c5-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a09c5-139">类型：</span><span class="sxs-lookup"><span data-stu-id="a09c5-139">Type:</span></span>

*   <span data-ttu-id="a09c5-140">String</span><span class="sxs-lookup"><span data-stu-id="a09c5-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a09c5-141">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-141">Requirements</span></span>

|<span data-ttu-id="a09c5-142">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-142">Requirement</span></span>| <span data-ttu-id="a09c5-143">值</span><span class="sxs-lookup"><span data-stu-id="a09c5-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="a09c5-144">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a09c5-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a09c5-145">1.0</span><span class="sxs-lookup"><span data-stu-id="a09c5-145">1.0</span></span>|
|[<span data-ttu-id="a09c5-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a09c5-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a09c5-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a09c5-147">ReadItem</span></span>|
|[<span data-ttu-id="a09c5-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a09c5-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a09c5-149">Compose or read</span><span class="sxs-lookup"><span data-stu-id="a09c5-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a09c5-150">示例</span><span class="sxs-lookup"><span data-stu-id="a09c5-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a09c5-151">timeZone :字符串</span><span class="sxs-lookup"><span data-stu-id="a09c5-151">timeZone :String</span></span>

<span data-ttu-id="a09c5-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="a09c5-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a09c5-153">类型：</span><span class="sxs-lookup"><span data-stu-id="a09c5-153">Type:</span></span>

*   <span data-ttu-id="a09c5-154">String</span><span class="sxs-lookup"><span data-stu-id="a09c5-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a09c5-155">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-155">Requirements</span></span>

|<span data-ttu-id="a09c5-156">要求</span><span class="sxs-lookup"><span data-stu-id="a09c5-156">Requirement</span></span>| <span data-ttu-id="a09c5-157">值</span><span class="sxs-lookup"><span data-stu-id="a09c5-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="a09c5-158">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a09c5-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a09c5-159">1.0</span><span class="sxs-lookup"><span data-stu-id="a09c5-159">1.0</span></span>|
|[<span data-ttu-id="a09c5-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a09c5-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a09c5-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a09c5-161">ReadItem</span></span>|
|[<span data-ttu-id="a09c5-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a09c5-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a09c5-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a09c5-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a09c5-164">示例</span><span class="sxs-lookup"><span data-stu-id="a09c5-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```