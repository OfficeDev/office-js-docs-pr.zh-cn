
# <a name="userprofile"></a><span data-ttu-id="5c1cf-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="5c1cf-101">userProfile</span></span>

### <span data-ttu-id="5c1cf-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="5c1cf-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c1cf-104">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-104">Requirements</span></span>

|<span data-ttu-id="5c1cf-105">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-105">Requirement</span></span>| <span data-ttu-id="5c1cf-106">值</span><span class="sxs-lookup"><span data-stu-id="5c1cf-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c1cf-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c1cf-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5c1cf-108">1.0</span></span>|
|[<span data-ttu-id="5c1cf-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c1cf-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c1cf-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c1cf-110">ReadItem</span></span>|
|[<span data-ttu-id="5c1cf-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c1cf-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c1cf-112">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="5c1cf-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5c1cf-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="5c1cf-113">Members and methods</span></span>

| <span data-ttu-id="5c1cf-114">成员</span><span class="sxs-lookup"><span data-stu-id="5c1cf-114">Member</span></span> | <span data-ttu-id="5c1cf-115">类型</span><span class="sxs-lookup"><span data-stu-id="5c1cf-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5c1cf-116">accountType</span><span class="sxs-lookup"><span data-stu-id="5c1cf-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="5c1cf-117">成员</span><span class="sxs-lookup"><span data-stu-id="5c1cf-117">Member</span></span> |
| [<span data-ttu-id="5c1cf-118">displayName</span><span class="sxs-lookup"><span data-stu-id="5c1cf-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="5c1cf-119">成员</span><span class="sxs-lookup"><span data-stu-id="5c1cf-119">Member</span></span> |
| [<span data-ttu-id="5c1cf-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5c1cf-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="5c1cf-121">成员</span><span class="sxs-lookup"><span data-stu-id="5c1cf-121">Member</span></span> |
| [<span data-ttu-id="5c1cf-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="5c1cf-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="5c1cf-123">成员</span><span class="sxs-lookup"><span data-stu-id="5c1cf-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="5c1cf-124">成员</span><span class="sxs-lookup"><span data-stu-id="5c1cf-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="5c1cf-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="5c1cf-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="5c1cf-126">当前仅 Outlook 2016  for Mac 或更高版本（内部版本 16.9.1212 或更高版本）支持此成员。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="5c1cf-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="5c1cf-128">下表列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="5c1cf-129">值</span><span class="sxs-lookup"><span data-stu-id="5c1cf-129">Value</span></span> | <span data-ttu-id="5c1cf-130">说明</span><span class="sxs-lookup"><span data-stu-id="5c1cf-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="5c1cf-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="5c1cf-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="5c1cf-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="5c1cf-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="5c1cf-135">类型：</span><span class="sxs-lookup"><span data-stu-id="5c1cf-135">Type:</span></span>

*   <span data-ttu-id="5c1cf-136">String</span><span class="sxs-lookup"><span data-stu-id="5c1cf-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c1cf-137">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-137">Requirements</span></span>

|<span data-ttu-id="5c1cf-138">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-138">Requirement</span></span>| <span data-ttu-id="5c1cf-139">值</span><span class="sxs-lookup"><span data-stu-id="5c1cf-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c1cf-140">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="5c1cf-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c1cf-141">1.6</span><span class="sxs-lookup"><span data-stu-id="5c1cf-141">-16</span></span> |
|[<span data-ttu-id="5c1cf-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c1cf-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c1cf-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c1cf-143">ReadItem</span></span>|
|[<span data-ttu-id="5c1cf-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c1cf-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c1cf-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c1cf-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c1cf-146">示例</span><span class="sxs-lookup"><span data-stu-id="5c1cf-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="5c1cf-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5c1cf-147">displayName :String</span></span>

<span data-ttu-id="5c1cf-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5c1cf-149">类型：</span><span class="sxs-lookup"><span data-stu-id="5c1cf-149">Type:</span></span>

*   <span data-ttu-id="5c1cf-150">字符串</span><span class="sxs-lookup"><span data-stu-id="5c1cf-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c1cf-151">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-151">Requirements</span></span>

|<span data-ttu-id="5c1cf-152">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-152">Requirement</span></span>| <span data-ttu-id="5c1cf-153">值</span><span class="sxs-lookup"><span data-stu-id="5c1cf-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c1cf-154">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c1cf-155">1.0</span><span class="sxs-lookup"><span data-stu-id="5c1cf-155">1.0</span></span>|
|[<span data-ttu-id="5c1cf-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c1cf-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c1cf-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c1cf-157">ReadItem</span></span>|
|[<span data-ttu-id="5c1cf-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c1cf-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c1cf-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c1cf-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c1cf-160">示例</span><span class="sxs-lookup"><span data-stu-id="5c1cf-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5c1cf-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5c1cf-161">emailAddress :String</span></span>

<span data-ttu-id="5c1cf-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5c1cf-163">类型：</span><span class="sxs-lookup"><span data-stu-id="5c1cf-163">Type:</span></span>

*   <span data-ttu-id="5c1cf-164">String</span><span class="sxs-lookup"><span data-stu-id="5c1cf-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c1cf-165">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-165">Requirements</span></span>

|<span data-ttu-id="5c1cf-166">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-166">Requirement</span></span>| <span data-ttu-id="5c1cf-167">值</span><span class="sxs-lookup"><span data-stu-id="5c1cf-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c1cf-168">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c1cf-169">1.0</span><span class="sxs-lookup"><span data-stu-id="5c1cf-169">1.0</span></span>|
|[<span data-ttu-id="5c1cf-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c1cf-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c1cf-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c1cf-171">ReadItem</span></span>|
|[<span data-ttu-id="5c1cf-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c1cf-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c1cf-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c1cf-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c1cf-174">示例</span><span class="sxs-lookup"><span data-stu-id="5c1cf-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5c1cf-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5c1cf-175">timeZone :String</span></span>

<span data-ttu-id="5c1cf-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="5c1cf-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5c1cf-177">类型：</span><span class="sxs-lookup"><span data-stu-id="5c1cf-177">Type:</span></span>

*   <span data-ttu-id="5c1cf-178">字符串</span><span class="sxs-lookup"><span data-stu-id="5c1cf-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c1cf-179">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-179">Requirements</span></span>

|<span data-ttu-id="5c1cf-180">要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-180">Requirement</span></span>| <span data-ttu-id="5c1cf-181">值</span><span class="sxs-lookup"><span data-stu-id="5c1cf-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c1cf-182">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="5c1cf-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c1cf-183">1.0</span><span class="sxs-lookup"><span data-stu-id="5c1cf-183">1.0</span></span>|
|[<span data-ttu-id="5c1cf-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c1cf-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c1cf-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c1cf-185">ReadItem</span></span>|
|[<span data-ttu-id="5c1cf-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c1cf-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c1cf-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c1cf-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c1cf-188">示例</span><span class="sxs-lookup"><span data-stu-id="5c1cf-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```