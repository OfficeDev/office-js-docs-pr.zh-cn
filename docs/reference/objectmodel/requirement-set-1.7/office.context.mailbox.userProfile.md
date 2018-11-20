
# <a name="userprofile"></a><span data-ttu-id="7894c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="7894c-101">userProfile</span></span>

### <span data-ttu-id="7894c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="7894c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="7894c-104">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-104">Requirements</span></span>

|<span data-ttu-id="7894c-105">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-105">Requirement</span></span>| <span data-ttu-id="7894c-106">值</span><span class="sxs-lookup"><span data-stu-id="7894c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="7894c-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7894c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7894c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="7894c-108">1.0</span></span>|
|[<span data-ttu-id="7894c-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7894c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7894c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7894c-110">ReadItem</span></span>|
|[<span data-ttu-id="7894c-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7894c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7894c-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7894c-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7894c-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="7894c-113">Members and methods</span></span>

| <span data-ttu-id="7894c-114">成员</span><span class="sxs-lookup"><span data-stu-id="7894c-114">Member</span></span> | <span data-ttu-id="7894c-115">类型</span><span class="sxs-lookup"><span data-stu-id="7894c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7894c-116">accountType</span><span class="sxs-lookup"><span data-stu-id="7894c-116">AccountType</span></span>](#accounttype-string) | <span data-ttu-id="7894c-117">成员</span><span class="sxs-lookup"><span data-stu-id="7894c-117">Member</span></span> |
| [<span data-ttu-id="7894c-118">displayName</span><span class="sxs-lookup"><span data-stu-id="7894c-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="7894c-119">成员</span><span class="sxs-lookup"><span data-stu-id="7894c-119">Member</span></span> |
| [<span data-ttu-id="7894c-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="7894c-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="7894c-121">成员</span><span class="sxs-lookup"><span data-stu-id="7894c-121">Member</span></span> |
| [<span data-ttu-id="7894c-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="7894c-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="7894c-123">成员</span><span class="sxs-lookup"><span data-stu-id="7894c-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="7894c-124">成员</span><span class="sxs-lookup"><span data-stu-id="7894c-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="7894c-125">accountType：字符串</span><span class="sxs-lookup"><span data-stu-id="7894c-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="7894c-126">此成员当前仅在 Outlook 2016 for Mac（内部版本 16.9.1212 或更高版本）中受支持。</span><span class="sxs-lookup"><span data-stu-id="7894c-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="7894c-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="7894c-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="7894c-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="7894c-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="7894c-129">值</span><span class="sxs-lookup"><span data-stu-id="7894c-129">Value</span></span> | <span data-ttu-id="7894c-130">说明</span><span class="sxs-lookup"><span data-stu-id="7894c-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="7894c-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="7894c-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="7894c-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="7894c-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="7894c-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="7894c-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="7894c-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="7894c-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="7894c-135">类型：</span><span class="sxs-lookup"><span data-stu-id="7894c-135">Type:</span></span>

*   <span data-ttu-id="7894c-136">String</span><span class="sxs-lookup"><span data-stu-id="7894c-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7894c-137">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-137">Requirements</span></span>

|<span data-ttu-id="7894c-138">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-138">Requirement</span></span>| <span data-ttu-id="7894c-139">值</span><span class="sxs-lookup"><span data-stu-id="7894c-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="7894c-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7894c-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7894c-141">1.6</span><span class="sxs-lookup"><span data-stu-id="7894c-141">-16</span></span> |
|[<span data-ttu-id="7894c-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7894c-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7894c-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7894c-143">ReadItem</span></span>|
|[<span data-ttu-id="7894c-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7894c-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7894c-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7894c-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7894c-146">示例</span><span class="sxs-lookup"><span data-stu-id="7894c-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="7894c-147">displayName：字符串</span><span class="sxs-lookup"><span data-stu-id="7894c-147">displayName :String</span></span>

<span data-ttu-id="7894c-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="7894c-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="7894c-149">类型：</span><span class="sxs-lookup"><span data-stu-id="7894c-149">Type:</span></span>

*   <span data-ttu-id="7894c-150">String</span><span class="sxs-lookup"><span data-stu-id="7894c-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7894c-151">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-151">Requirements</span></span>

|<span data-ttu-id="7894c-152">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-152">Requirement</span></span>| <span data-ttu-id="7894c-153">值</span><span class="sxs-lookup"><span data-stu-id="7894c-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="7894c-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7894c-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7894c-155">1.0</span><span class="sxs-lookup"><span data-stu-id="7894c-155">1.0</span></span>|
|[<span data-ttu-id="7894c-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7894c-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7894c-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7894c-157">ReadItem</span></span>|
|[<span data-ttu-id="7894c-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7894c-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7894c-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7894c-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7894c-160">示例</span><span class="sxs-lookup"><span data-stu-id="7894c-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="7894c-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="7894c-161">emailAddress :String</span></span>

<span data-ttu-id="7894c-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="7894c-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="7894c-163">类型：</span><span class="sxs-lookup"><span data-stu-id="7894c-163">Type:</span></span>

*   <span data-ttu-id="7894c-164">String</span><span class="sxs-lookup"><span data-stu-id="7894c-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7894c-165">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-165">Requirements</span></span>

|<span data-ttu-id="7894c-166">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-166">Requirement</span></span>| <span data-ttu-id="7894c-167">值</span><span class="sxs-lookup"><span data-stu-id="7894c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="7894c-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7894c-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7894c-169">1.0</span><span class="sxs-lookup"><span data-stu-id="7894c-169">1.0</span></span>|
|[<span data-ttu-id="7894c-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7894c-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7894c-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7894c-171">ReadItem</span></span>|
|[<span data-ttu-id="7894c-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7894c-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7894c-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7894c-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7894c-174">示例</span><span class="sxs-lookup"><span data-stu-id="7894c-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="7894c-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="7894c-175">timeZone :String</span></span>

<span data-ttu-id="7894c-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="7894c-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="7894c-177">类型：</span><span class="sxs-lookup"><span data-stu-id="7894c-177">Type:</span></span>

*   <span data-ttu-id="7894c-178">String</span><span class="sxs-lookup"><span data-stu-id="7894c-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7894c-179">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-179">Requirements</span></span>

|<span data-ttu-id="7894c-180">要求</span><span class="sxs-lookup"><span data-stu-id="7894c-180">Requirement</span></span>| <span data-ttu-id="7894c-181">值</span><span class="sxs-lookup"><span data-stu-id="7894c-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="7894c-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7894c-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7894c-183">1.0</span><span class="sxs-lookup"><span data-stu-id="7894c-183">1.0</span></span>|
|[<span data-ttu-id="7894c-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="7894c-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7894c-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7894c-185">ReadItem</span></span>|
|[<span data-ttu-id="7894c-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7894c-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7894c-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7894c-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7894c-188">示例</span><span class="sxs-lookup"><span data-stu-id="7894c-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```