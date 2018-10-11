# <a name="userprofile"></a><span data-ttu-id="c4569-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c4569-101">userProfile</span></span>

### <span data-ttu-id="c4569-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c4569-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4569-104">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-104">Requirements</span></span>

|<span data-ttu-id="c4569-105">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-105">Requirement</span></span>| <span data-ttu-id="c4569-106">值</span><span class="sxs-lookup"><span data-stu-id="c4569-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4569-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="c4569-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4569-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c4569-108">1.0</span></span>|
|[<span data-ttu-id="c4569-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4569-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4569-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4569-110">ReadItem</span></span>|
|[<span data-ttu-id="c4569-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4569-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4569-112">Compose or read</span><span class="sxs-lookup"><span data-stu-id="c4569-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c4569-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c4569-113">Members and methods</span></span>

| <span data-ttu-id="c4569-114">Member</span><span class="sxs-lookup"><span data-stu-id="c4569-114">Member</span></span> | <span data-ttu-id="c4569-115">类型</span><span class="sxs-lookup"><span data-stu-id="c4569-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c4569-116">displayName</span><span class="sxs-lookup"><span data-stu-id="c4569-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="c4569-117">Member</span><span class="sxs-lookup"><span data-stu-id="c4569-117">Member</span></span> |
| [<span data-ttu-id="c4569-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c4569-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c4569-119">Member</span><span class="sxs-lookup"><span data-stu-id="c4569-119">Member</span></span> |
| [<span data-ttu-id="c4569-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="c4569-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c4569-121">Member</span><span class="sxs-lookup"><span data-stu-id="c4569-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c4569-122">成员</span><span class="sxs-lookup"><span data-stu-id="c4569-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c4569-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c4569-123">displayName :String</span></span>

<span data-ttu-id="c4569-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="c4569-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c4569-125">类型：</span><span class="sxs-lookup"><span data-stu-id="c4569-125">Type:</span></span>

*   <span data-ttu-id="c4569-126">String</span><span class="sxs-lookup"><span data-stu-id="c4569-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4569-127">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-127">Requirements</span></span>

|<span data-ttu-id="c4569-128">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-128">Requirement</span></span>| <span data-ttu-id="c4569-129">值</span><span class="sxs-lookup"><span data-stu-id="c4569-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4569-130">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="c4569-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4569-131">1.0</span><span class="sxs-lookup"><span data-stu-id="c4569-131">1.0</span></span>|
|[<span data-ttu-id="c4569-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4569-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4569-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4569-133">ReadItem</span></span>|
|[<span data-ttu-id="c4569-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4569-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4569-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4569-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4569-136">示例</span><span class="sxs-lookup"><span data-stu-id="c4569-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c4569-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c4569-137">emailAddress :String</span></span>

<span data-ttu-id="c4569-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c4569-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c4569-139">类型：</span><span class="sxs-lookup"><span data-stu-id="c4569-139">Type:</span></span>

*   <span data-ttu-id="c4569-140">String</span><span class="sxs-lookup"><span data-stu-id="c4569-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4569-141">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-141">Requirements</span></span>

|<span data-ttu-id="c4569-142">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-142">Requirement</span></span>| <span data-ttu-id="c4569-143">值</span><span class="sxs-lookup"><span data-stu-id="c4569-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4569-144">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="c4569-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4569-145">1.0</span><span class="sxs-lookup"><span data-stu-id="c4569-145">1.0</span></span>|
|[<span data-ttu-id="c4569-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4569-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4569-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4569-147">ReadItem</span></span>|
|[<span data-ttu-id="c4569-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4569-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4569-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4569-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4569-150">示例</span><span class="sxs-lookup"><span data-stu-id="c4569-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c4569-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c4569-151">timeZone :String</span></span>

<span data-ttu-id="c4569-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="c4569-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c4569-153">类型：</span><span class="sxs-lookup"><span data-stu-id="c4569-153">Type:</span></span>

*   <span data-ttu-id="c4569-154">String</span><span class="sxs-lookup"><span data-stu-id="c4569-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4569-155">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-155">Requirements</span></span>

|<span data-ttu-id="c4569-156">要求</span><span class="sxs-lookup"><span data-stu-id="c4569-156">Requirement</span></span>| <span data-ttu-id="c4569-157">值</span><span class="sxs-lookup"><span data-stu-id="c4569-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4569-158">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="c4569-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4569-159">1.0</span><span class="sxs-lookup"><span data-stu-id="c4569-159">1.0</span></span>|
|[<span data-ttu-id="c4569-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4569-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4569-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4569-161">ReadItem</span></span>|
|[<span data-ttu-id="c4569-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4569-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4569-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4569-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4569-164">示例</span><span class="sxs-lookup"><span data-stu-id="c4569-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```