---
title: Office.context.mailbox.userProfile - 预览要求集
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 061ee8367005f4af0795c4d9e1236d0b2443521a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432814"
---
# <a name="userprofile"></a><span data-ttu-id="c7bbc-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c7bbc-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c7bbc-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c7bbc-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7bbc-104">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-104">Requirements</span></span>

|<span data-ttu-id="c7bbc-105">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-105">Requirement</span></span>| <span data-ttu-id="c7bbc-106">值</span><span class="sxs-lookup"><span data-stu-id="c7bbc-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7bbc-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c7bbc-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7bbc-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c7bbc-108">1.0</span></span>|
|[<span data-ttu-id="c7bbc-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7bbc-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7bbc-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7bbc-110">ReadItem</span></span>|
|[<span data-ttu-id="c7bbc-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7bbc-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7bbc-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7bbc-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c7bbc-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c7bbc-113">Members and methods</span></span>

| <span data-ttu-id="c7bbc-114">成员</span><span class="sxs-lookup"><span data-stu-id="c7bbc-114">Member</span></span> | <span data-ttu-id="c7bbc-115">类型</span><span class="sxs-lookup"><span data-stu-id="c7bbc-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c7bbc-116">accountType</span><span class="sxs-lookup"><span data-stu-id="c7bbc-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="c7bbc-117">成员</span><span class="sxs-lookup"><span data-stu-id="c7bbc-117">Member</span></span> |
| [<span data-ttu-id="c7bbc-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c7bbc-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c7bbc-119">成员</span><span class="sxs-lookup"><span data-stu-id="c7bbc-119">Member</span></span> |
| [<span data-ttu-id="c7bbc-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c7bbc-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c7bbc-121">成员</span><span class="sxs-lookup"><span data-stu-id="c7bbc-121">Member</span></span> |
| [<span data-ttu-id="c7bbc-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c7bbc-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c7bbc-123">成员</span><span class="sxs-lookup"><span data-stu-id="c7bbc-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c7bbc-124">成员</span><span class="sxs-lookup"><span data-stu-id="c7bbc-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c7bbc-125">accountType：字符串</span><span class="sxs-lookup"><span data-stu-id="c7bbc-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c7bbc-126">此成员当前仅在适用于 Mac 的 Outlook 2016 或更高版本（内部版本 16.9.1212 或更高版本）中受支持。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="c7bbc-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c7bbc-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="c7bbc-129">值</span><span class="sxs-lookup"><span data-stu-id="c7bbc-129">Value</span></span> | <span data-ttu-id="c7bbc-130">描述</span><span class="sxs-lookup"><span data-stu-id="c7bbc-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c7bbc-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c7bbc-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c7bbc-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c7bbc-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c7bbc-135">类型：</span><span class="sxs-lookup"><span data-stu-id="c7bbc-135">Type:</span></span>

*   <span data-ttu-id="c7bbc-136">String</span><span class="sxs-lookup"><span data-stu-id="c7bbc-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7bbc-137">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-137">Requirements</span></span>

|<span data-ttu-id="c7bbc-138">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-138">Requirement</span></span>| <span data-ttu-id="c7bbc-139">值</span><span class="sxs-lookup"><span data-stu-id="c7bbc-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7bbc-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c7bbc-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7bbc-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c7bbc-141">1.6</span></span> |
|[<span data-ttu-id="c7bbc-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7bbc-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7bbc-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7bbc-143">ReadItem</span></span>|
|[<span data-ttu-id="c7bbc-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7bbc-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7bbc-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7bbc-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7bbc-146">示例</span><span class="sxs-lookup"><span data-stu-id="c7bbc-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="c7bbc-147">displayName：字符串</span><span class="sxs-lookup"><span data-stu-id="c7bbc-147">displayName :String</span></span>

<span data-ttu-id="c7bbc-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c7bbc-149">类型：</span><span class="sxs-lookup"><span data-stu-id="c7bbc-149">Type:</span></span>

*   <span data-ttu-id="c7bbc-150">String</span><span class="sxs-lookup"><span data-stu-id="c7bbc-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7bbc-151">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-151">Requirements</span></span>

|<span data-ttu-id="c7bbc-152">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-152">Requirement</span></span>| <span data-ttu-id="c7bbc-153">值</span><span class="sxs-lookup"><span data-stu-id="c7bbc-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7bbc-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c7bbc-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7bbc-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c7bbc-155">1.0</span></span>|
|[<span data-ttu-id="c7bbc-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7bbc-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7bbc-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7bbc-157">ReadItem</span></span>|
|[<span data-ttu-id="c7bbc-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7bbc-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7bbc-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7bbc-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7bbc-160">示例</span><span class="sxs-lookup"><span data-stu-id="c7bbc-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c7bbc-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c7bbc-161">emailAddress :String</span></span>

<span data-ttu-id="c7bbc-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c7bbc-163">类型：</span><span class="sxs-lookup"><span data-stu-id="c7bbc-163">Type:</span></span>

*   <span data-ttu-id="c7bbc-164">String</span><span class="sxs-lookup"><span data-stu-id="c7bbc-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7bbc-165">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-165">Requirements</span></span>

|<span data-ttu-id="c7bbc-166">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-166">Requirement</span></span>| <span data-ttu-id="c7bbc-167">值</span><span class="sxs-lookup"><span data-stu-id="c7bbc-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7bbc-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c7bbc-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7bbc-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c7bbc-169">1.0</span></span>|
|[<span data-ttu-id="c7bbc-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7bbc-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7bbc-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7bbc-171">ReadItem</span></span>|
|[<span data-ttu-id="c7bbc-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7bbc-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7bbc-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7bbc-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7bbc-174">示例</span><span class="sxs-lookup"><span data-stu-id="c7bbc-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c7bbc-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c7bbc-175">timeZone :String</span></span>

<span data-ttu-id="c7bbc-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="c7bbc-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c7bbc-177">类型：</span><span class="sxs-lookup"><span data-stu-id="c7bbc-177">Type:</span></span>

*   <span data-ttu-id="c7bbc-178">String</span><span class="sxs-lookup"><span data-stu-id="c7bbc-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7bbc-179">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-179">Requirements</span></span>

|<span data-ttu-id="c7bbc-180">要求</span><span class="sxs-lookup"><span data-stu-id="c7bbc-180">Requirement</span></span>| <span data-ttu-id="c7bbc-181">值</span><span class="sxs-lookup"><span data-stu-id="c7bbc-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7bbc-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c7bbc-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7bbc-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c7bbc-183">1.0</span></span>|
|[<span data-ttu-id="c7bbc-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7bbc-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7bbc-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7bbc-185">ReadItem</span></span>|
|[<span data-ttu-id="c7bbc-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7bbc-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7bbc-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7bbc-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7bbc-188">示例</span><span class="sxs-lookup"><span data-stu-id="c7bbc-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```