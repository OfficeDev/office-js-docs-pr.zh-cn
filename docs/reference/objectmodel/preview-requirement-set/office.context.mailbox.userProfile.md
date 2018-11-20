
# <a name="userprofile"></a><span data-ttu-id="c1a87-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c1a87-101">userProfile</span></span>

### <span data-ttu-id="c1a87-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c1a87-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1a87-104">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-104">Requirements</span></span>

|<span data-ttu-id="c1a87-105">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-105">Requirement</span></span>| <span data-ttu-id="c1a87-106">值</span><span class="sxs-lookup"><span data-stu-id="c1a87-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1a87-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c1a87-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1a87-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c1a87-108">1.0</span></span>|
|[<span data-ttu-id="c1a87-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c1a87-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1a87-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1a87-110">ReadItem</span></span>|
|[<span data-ttu-id="c1a87-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c1a87-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1a87-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c1a87-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1a87-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c1a87-113">Members and methods</span></span>

| <span data-ttu-id="c1a87-114">成员</span><span class="sxs-lookup"><span data-stu-id="c1a87-114">Member</span></span> | <span data-ttu-id="c1a87-115">类型</span><span class="sxs-lookup"><span data-stu-id="c1a87-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1a87-116">accountType</span><span class="sxs-lookup"><span data-stu-id="c1a87-116">AccountType</span></span>](#accounttype-string) | <span data-ttu-id="c1a87-117">成员</span><span class="sxs-lookup"><span data-stu-id="c1a87-117">Member</span></span> |
| [<span data-ttu-id="c1a87-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c1a87-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c1a87-119">成员</span><span class="sxs-lookup"><span data-stu-id="c1a87-119">Member</span></span> |
| [<span data-ttu-id="c1a87-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c1a87-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c1a87-121">成员</span><span class="sxs-lookup"><span data-stu-id="c1a87-121">Member</span></span> |
| [<span data-ttu-id="c1a87-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c1a87-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c1a87-123">成员</span><span class="sxs-lookup"><span data-stu-id="c1a87-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c1a87-124">成员</span><span class="sxs-lookup"><span data-stu-id="c1a87-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c1a87-125">accountType：字符串</span><span class="sxs-lookup"><span data-stu-id="c1a87-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c1a87-126">此成员当前仅在适用于 Mac 的 Outlook 2016 或更高版本（内部版本 16.9.1212 或更高版本）中受支持。</span><span class="sxs-lookup"><span data-stu-id="c1a87-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="c1a87-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="c1a87-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c1a87-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="c1a87-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="c1a87-129">值</span><span class="sxs-lookup"><span data-stu-id="c1a87-129">Value</span></span> | <span data-ttu-id="c1a87-130">说明</span><span class="sxs-lookup"><span data-stu-id="c1a87-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c1a87-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="c1a87-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c1a87-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="c1a87-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c1a87-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="c1a87-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c1a87-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="c1a87-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c1a87-135">类型：</span><span class="sxs-lookup"><span data-stu-id="c1a87-135">Type:</span></span>

*   <span data-ttu-id="c1a87-136">String</span><span class="sxs-lookup"><span data-stu-id="c1a87-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1a87-137">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-137">Requirements</span></span>

|<span data-ttu-id="c1a87-138">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-138">Requirement</span></span>| <span data-ttu-id="c1a87-139">值</span><span class="sxs-lookup"><span data-stu-id="c1a87-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1a87-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c1a87-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1a87-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c1a87-141">-16</span></span> |
|[<span data-ttu-id="c1a87-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c1a87-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1a87-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1a87-143">ReadItem</span></span>|
|[<span data-ttu-id="c1a87-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c1a87-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1a87-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c1a87-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1a87-146">示例</span><span class="sxs-lookup"><span data-stu-id="c1a87-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="c1a87-147">displayName：字符串</span><span class="sxs-lookup"><span data-stu-id="c1a87-147">displayName :String</span></span>

<span data-ttu-id="c1a87-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="c1a87-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c1a87-149">类型：</span><span class="sxs-lookup"><span data-stu-id="c1a87-149">Type:</span></span>

*   <span data-ttu-id="c1a87-150">String</span><span class="sxs-lookup"><span data-stu-id="c1a87-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1a87-151">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-151">Requirements</span></span>

|<span data-ttu-id="c1a87-152">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-152">Requirement</span></span>| <span data-ttu-id="c1a87-153">值</span><span class="sxs-lookup"><span data-stu-id="c1a87-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1a87-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c1a87-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1a87-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c1a87-155">1.0</span></span>|
|[<span data-ttu-id="c1a87-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c1a87-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1a87-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1a87-157">ReadItem</span></span>|
|[<span data-ttu-id="c1a87-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c1a87-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1a87-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c1a87-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1a87-160">示例</span><span class="sxs-lookup"><span data-stu-id="c1a87-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c1a87-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c1a87-161">emailAddress :String</span></span>

<span data-ttu-id="c1a87-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c1a87-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c1a87-163">类型：</span><span class="sxs-lookup"><span data-stu-id="c1a87-163">Type:</span></span>

*   <span data-ttu-id="c1a87-164">String</span><span class="sxs-lookup"><span data-stu-id="c1a87-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1a87-165">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-165">Requirements</span></span>

|<span data-ttu-id="c1a87-166">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-166">Requirement</span></span>| <span data-ttu-id="c1a87-167">值</span><span class="sxs-lookup"><span data-stu-id="c1a87-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1a87-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c1a87-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1a87-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c1a87-169">1.0</span></span>|
|[<span data-ttu-id="c1a87-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c1a87-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1a87-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1a87-171">ReadItem</span></span>|
|[<span data-ttu-id="c1a87-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c1a87-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1a87-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c1a87-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1a87-174">示例</span><span class="sxs-lookup"><span data-stu-id="c1a87-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c1a87-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c1a87-175">timeZone :String</span></span>

<span data-ttu-id="c1a87-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="c1a87-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c1a87-177">类型：</span><span class="sxs-lookup"><span data-stu-id="c1a87-177">Type:</span></span>

*   <span data-ttu-id="c1a87-178">String</span><span class="sxs-lookup"><span data-stu-id="c1a87-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1a87-179">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-179">Requirements</span></span>

|<span data-ttu-id="c1a87-180">要求</span><span class="sxs-lookup"><span data-stu-id="c1a87-180">Requirement</span></span>| <span data-ttu-id="c1a87-181">值</span><span class="sxs-lookup"><span data-stu-id="c1a87-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1a87-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c1a87-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1a87-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c1a87-183">1.0</span></span>|
|[<span data-ttu-id="c1a87-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c1a87-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1a87-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1a87-185">ReadItem</span></span>|
|[<span data-ttu-id="c1a87-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c1a87-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1a87-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c1a87-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1a87-188">示例</span><span class="sxs-lookup"><span data-stu-id="c1a87-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```