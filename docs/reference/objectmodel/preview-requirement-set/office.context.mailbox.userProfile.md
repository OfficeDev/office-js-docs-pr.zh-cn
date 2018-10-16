
# <a name="userprofile"></a><span data-ttu-id="36eea-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="36eea-101">userProfile</span></span>

### <span data-ttu-id="36eea-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="36eea-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="36eea-104">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-104">Requirements</span></span>

|<span data-ttu-id="36eea-105">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-105">Requirement</span></span>| <span data-ttu-id="36eea-106">值</span><span class="sxs-lookup"><span data-stu-id="36eea-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="36eea-107">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="36eea-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36eea-108">1.0</span><span class="sxs-lookup"><span data-stu-id="36eea-108">1.0</span></span>|
|[<span data-ttu-id="36eea-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="36eea-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="36eea-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="36eea-110">ReadItem</span></span>|
|[<span data-ttu-id="36eea-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36eea-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36eea-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36eea-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="36eea-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="36eea-113">Members and methods</span></span>

| <span data-ttu-id="36eea-114">成员</span><span class="sxs-lookup"><span data-stu-id="36eea-114">Member</span></span> | <span data-ttu-id="36eea-115">类型</span><span class="sxs-lookup"><span data-stu-id="36eea-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="36eea-116">accountType</span><span class="sxs-lookup"><span data-stu-id="36eea-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="36eea-117">成员</span><span class="sxs-lookup"><span data-stu-id="36eea-117">Member</span></span> |
| [<span data-ttu-id="36eea-118">displayName</span><span class="sxs-lookup"><span data-stu-id="36eea-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="36eea-119">成员</span><span class="sxs-lookup"><span data-stu-id="36eea-119">Member</span></span> |
| [<span data-ttu-id="36eea-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="36eea-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="36eea-121">成员</span><span class="sxs-lookup"><span data-stu-id="36eea-121">Member</span></span> |
| [<span data-ttu-id="36eea-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="36eea-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="36eea-123">成员</span><span class="sxs-lookup"><span data-stu-id="36eea-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="36eea-124">成员</span><span class="sxs-lookup"><span data-stu-id="36eea-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="36eea-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="36eea-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="36eea-126">当前仅 Outlook 2016  for Mac 或更高版本（内部版本 16.9.1212 或更高版本）支持此成员。</span><span class="sxs-lookup"><span data-stu-id="36eea-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="36eea-p102">获取与邮箱关联用户的帐户类型。下表列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="36eea-p102">Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.</span></span>

| <span data-ttu-id="36eea-129">值</span><span class="sxs-lookup"><span data-stu-id="36eea-129">Value</span></span> | <span data-ttu-id="36eea-130">说明</span><span class="sxs-lookup"><span data-stu-id="36eea-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="36eea-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="36eea-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="36eea-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="36eea-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="36eea-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="36eea-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="36eea-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="36eea-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="36eea-135">类型：</span><span class="sxs-lookup"><span data-stu-id="36eea-135">Type:</span></span>

*   <span data-ttu-id="36eea-136">String</span><span class="sxs-lookup"><span data-stu-id="36eea-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36eea-137">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-137">Requirements</span></span>

|<span data-ttu-id="36eea-138">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-138">Requirement</span></span>| <span data-ttu-id="36eea-139">值</span><span class="sxs-lookup"><span data-stu-id="36eea-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="36eea-140">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="36eea-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36eea-141">1.6</span><span class="sxs-lookup"><span data-stu-id="36eea-141">-16</span></span> |
|[<span data-ttu-id="36eea-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="36eea-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="36eea-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="36eea-143">ReadItem</span></span>|
|[<span data-ttu-id="36eea-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36eea-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36eea-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36eea-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="36eea-146">示例</span><span class="sxs-lookup"><span data-stu-id="36eea-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="36eea-147">displayName :字符串</span><span class="sxs-lookup"><span data-stu-id="36eea-147">displayName :String</span></span>

<span data-ttu-id="36eea-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="36eea-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="36eea-149">类型：</span><span class="sxs-lookup"><span data-stu-id="36eea-149">Type:</span></span>

*   <span data-ttu-id="36eea-150">String</span><span class="sxs-lookup"><span data-stu-id="36eea-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36eea-151">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-151">Requirements</span></span>

|<span data-ttu-id="36eea-152">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-152">Requirement</span></span>| <span data-ttu-id="36eea-153">值</span><span class="sxs-lookup"><span data-stu-id="36eea-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="36eea-154">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="36eea-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36eea-155">1.0</span><span class="sxs-lookup"><span data-stu-id="36eea-155">1.0</span></span>|
|[<span data-ttu-id="36eea-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="36eea-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="36eea-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="36eea-157">ReadItem</span></span>|
|[<span data-ttu-id="36eea-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36eea-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36eea-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36eea-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="36eea-160">示例</span><span class="sxs-lookup"><span data-stu-id="36eea-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="36eea-161">emailAddress :字符串</span><span class="sxs-lookup"><span data-stu-id="36eea-161">emailAddress :String</span></span>

<span data-ttu-id="36eea-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="36eea-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="36eea-163">类型：</span><span class="sxs-lookup"><span data-stu-id="36eea-163">Type:</span></span>

*   <span data-ttu-id="36eea-164">String</span><span class="sxs-lookup"><span data-stu-id="36eea-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36eea-165">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-165">Requirements</span></span>

|<span data-ttu-id="36eea-166">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-166">Requirement</span></span>| <span data-ttu-id="36eea-167">值</span><span class="sxs-lookup"><span data-stu-id="36eea-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="36eea-168">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="36eea-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36eea-169">1.0</span><span class="sxs-lookup"><span data-stu-id="36eea-169">1.0</span></span>|
|[<span data-ttu-id="36eea-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="36eea-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="36eea-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="36eea-171">ReadItem</span></span>|
|[<span data-ttu-id="36eea-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36eea-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36eea-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36eea-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="36eea-174">示例</span><span class="sxs-lookup"><span data-stu-id="36eea-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="36eea-175">timeZone :字符串</span><span class="sxs-lookup"><span data-stu-id="36eea-175">timeZone :String</span></span>

<span data-ttu-id="36eea-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="36eea-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="36eea-177">类型：</span><span class="sxs-lookup"><span data-stu-id="36eea-177">Type:</span></span>

*   <span data-ttu-id="36eea-178">String</span><span class="sxs-lookup"><span data-stu-id="36eea-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36eea-179">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-179">Requirements</span></span>

|<span data-ttu-id="36eea-180">要求</span><span class="sxs-lookup"><span data-stu-id="36eea-180">Requirement</span></span>| <span data-ttu-id="36eea-181">值</span><span class="sxs-lookup"><span data-stu-id="36eea-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="36eea-182">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="36eea-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36eea-183">1.0</span><span class="sxs-lookup"><span data-stu-id="36eea-183">1.0</span></span>|
|[<span data-ttu-id="36eea-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="36eea-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="36eea-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="36eea-185">ReadItem</span></span>|
|[<span data-ttu-id="36eea-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36eea-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36eea-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36eea-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="36eea-188">示例</span><span class="sxs-lookup"><span data-stu-id="36eea-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```