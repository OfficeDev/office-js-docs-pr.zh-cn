# <a name="userprofile"></a><span data-ttu-id="9c1e6-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="9c1e6-101">userProfile</span></span>

### <span data-ttu-id="9c1e6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="9c1e6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c1e6-104">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-104">Requirements</span></span>

|<span data-ttu-id="9c1e6-105">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-105">Requirement</span></span>| <span data-ttu-id="9c1e6-106">值</span><span class="sxs-lookup"><span data-stu-id="9c1e6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c1e6-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9c1e6-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c1e6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9c1e6-108">1.0</span></span>|
|[<span data-ttu-id="9c1e6-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9c1e6-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c1e6-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c1e6-110">ReadItem</span></span>|
|[<span data-ttu-id="9c1e6-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9c1e6-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c1e6-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9c1e6-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9c1e6-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="9c1e6-113">Members and methods</span></span>

| <span data-ttu-id="9c1e6-114">成员</span><span class="sxs-lookup"><span data-stu-id="9c1e6-114">Member</span></span> | <span data-ttu-id="9c1e6-115">类型</span><span class="sxs-lookup"><span data-stu-id="9c1e6-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9c1e6-116">displayName</span><span class="sxs-lookup"><span data-stu-id="9c1e6-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="9c1e6-117">成员</span><span class="sxs-lookup"><span data-stu-id="9c1e6-117">Member</span></span> |
| [<span data-ttu-id="9c1e6-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="9c1e6-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="9c1e6-119">成员</span><span class="sxs-lookup"><span data-stu-id="9c1e6-119">Member</span></span> |
| [<span data-ttu-id="9c1e6-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="9c1e6-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="9c1e6-121">成员</span><span class="sxs-lookup"><span data-stu-id="9c1e6-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="9c1e6-122">成员</span><span class="sxs-lookup"><span data-stu-id="9c1e6-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="9c1e6-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="9c1e6-123">displayName :String</span></span>

<span data-ttu-id="9c1e6-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="9c1e6-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="9c1e6-125">类型：</span><span class="sxs-lookup"><span data-stu-id="9c1e6-125">Type:</span></span>

*   <span data-ttu-id="9c1e6-126">String</span><span class="sxs-lookup"><span data-stu-id="9c1e6-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c1e6-127">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-127">Requirements</span></span>

|<span data-ttu-id="9c1e6-128">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-128">Requirement</span></span>| <span data-ttu-id="9c1e6-129">值</span><span class="sxs-lookup"><span data-stu-id="9c1e6-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c1e6-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9c1e6-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c1e6-131">1.0</span><span class="sxs-lookup"><span data-stu-id="9c1e6-131">1.0</span></span>|
|[<span data-ttu-id="9c1e6-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9c1e6-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c1e6-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c1e6-133">ReadItem</span></span>|
|[<span data-ttu-id="9c1e6-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9c1e6-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c1e6-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9c1e6-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c1e6-136">示例</span><span class="sxs-lookup"><span data-stu-id="9c1e6-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="9c1e6-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="9c1e6-137">emailAddress :String</span></span>

<span data-ttu-id="9c1e6-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="9c1e6-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="9c1e6-139">类型：</span><span class="sxs-lookup"><span data-stu-id="9c1e6-139">Type:</span></span>

*   <span data-ttu-id="9c1e6-140">String</span><span class="sxs-lookup"><span data-stu-id="9c1e6-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c1e6-141">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-141">Requirements</span></span>

|<span data-ttu-id="9c1e6-142">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-142">Requirement</span></span>| <span data-ttu-id="9c1e6-143">值</span><span class="sxs-lookup"><span data-stu-id="9c1e6-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c1e6-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9c1e6-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c1e6-145">1.0</span><span class="sxs-lookup"><span data-stu-id="9c1e6-145">1.0</span></span>|
|[<span data-ttu-id="9c1e6-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9c1e6-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c1e6-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c1e6-147">ReadItem</span></span>|
|[<span data-ttu-id="9c1e6-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9c1e6-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c1e6-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9c1e6-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c1e6-150">示例</span><span class="sxs-lookup"><span data-stu-id="9c1e6-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="9c1e6-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="9c1e6-151">timeZone :String</span></span>

<span data-ttu-id="9c1e6-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="9c1e6-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="9c1e6-153">类型：</span><span class="sxs-lookup"><span data-stu-id="9c1e6-153">Type:</span></span>

*   <span data-ttu-id="9c1e6-154">String</span><span class="sxs-lookup"><span data-stu-id="9c1e6-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c1e6-155">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-155">Requirements</span></span>

|<span data-ttu-id="9c1e6-156">要求</span><span class="sxs-lookup"><span data-stu-id="9c1e6-156">Requirement</span></span>| <span data-ttu-id="9c1e6-157">值</span><span class="sxs-lookup"><span data-stu-id="9c1e6-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c1e6-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9c1e6-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c1e6-159">1.0</span><span class="sxs-lookup"><span data-stu-id="9c1e6-159">1.0</span></span>|
|[<span data-ttu-id="9c1e6-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9c1e6-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c1e6-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c1e6-161">ReadItem</span></span>|
|[<span data-ttu-id="9c1e6-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9c1e6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c1e6-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9c1e6-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c1e6-164">示例</span><span class="sxs-lookup"><span data-stu-id="9c1e6-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```