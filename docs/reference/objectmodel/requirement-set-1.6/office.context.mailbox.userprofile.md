---
title: Office.context.mailbox.userProfile - 要求集 1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 09457a41fe68ae03e035d3d3f4b80b139be348e0
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067872"
---
# <a name="userprofile"></a><span data-ttu-id="ca533-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ca533-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ca533-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ca533-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca533-104">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-104">Requirements</span></span>

|<span data-ttu-id="ca533-105">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-105">Requirement</span></span>| <span data-ttu-id="ca533-106">值</span><span class="sxs-lookup"><span data-stu-id="ca533-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca533-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca533-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca533-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ca533-108">1.0</span></span>|
|[<span data-ttu-id="ca533-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca533-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca533-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca533-110">ReadItem</span></span>|
|[<span data-ttu-id="ca533-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca533-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca533-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca533-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ca533-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="ca533-113">Members and methods</span></span>

| <span data-ttu-id="ca533-114">成员</span><span class="sxs-lookup"><span data-stu-id="ca533-114">Member</span></span> | <span data-ttu-id="ca533-115">类型</span><span class="sxs-lookup"><span data-stu-id="ca533-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ca533-116">accountType</span><span class="sxs-lookup"><span data-stu-id="ca533-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="ca533-117">成员</span><span class="sxs-lookup"><span data-stu-id="ca533-117">Member</span></span> |
| [<span data-ttu-id="ca533-118">displayName</span><span class="sxs-lookup"><span data-stu-id="ca533-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="ca533-119">成员</span><span class="sxs-lookup"><span data-stu-id="ca533-119">Member</span></span> |
| [<span data-ttu-id="ca533-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ca533-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ca533-121">成员</span><span class="sxs-lookup"><span data-stu-id="ca533-121">Member</span></span> |
| [<span data-ttu-id="ca533-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="ca533-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ca533-123">Member</span><span class="sxs-lookup"><span data-stu-id="ca533-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ca533-124">成员</span><span class="sxs-lookup"><span data-stu-id="ca533-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="ca533-125">accountType：字符串</span><span class="sxs-lookup"><span data-stu-id="ca533-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="ca533-126">此成员当前仅在适用于 Mac 的 Outlook 2016 或更高版本（内部版本 16.9.1212 或更高版本）中受支持。</span><span class="sxs-lookup"><span data-stu-id="ca533-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="ca533-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="ca533-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="ca533-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="ca533-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="ca533-129">值</span><span class="sxs-lookup"><span data-stu-id="ca533-129">Value</span></span> | <span data-ttu-id="ca533-130">说明</span><span class="sxs-lookup"><span data-stu-id="ca533-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="ca533-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="ca533-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="ca533-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="ca533-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="ca533-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="ca533-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="ca533-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="ca533-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="ca533-135">Type</span><span class="sxs-lookup"><span data-stu-id="ca533-135">Type</span></span>

*   <span data-ttu-id="ca533-136">String</span><span class="sxs-lookup"><span data-stu-id="ca533-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca533-137">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-137">Requirements</span></span>

|<span data-ttu-id="ca533-138">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-138">Requirement</span></span>| <span data-ttu-id="ca533-139">值</span><span class="sxs-lookup"><span data-stu-id="ca533-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca533-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca533-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca533-141">1.6</span><span class="sxs-lookup"><span data-stu-id="ca533-141">1.6</span></span> |
|[<span data-ttu-id="ca533-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca533-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca533-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca533-143">ReadItem</span></span>|
|[<span data-ttu-id="ca533-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca533-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca533-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca533-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca533-146">示例</span><span class="sxs-lookup"><span data-stu-id="ca533-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="ca533-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ca533-147">displayName :String</span></span>

<span data-ttu-id="ca533-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="ca533-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ca533-149">Type</span><span class="sxs-lookup"><span data-stu-id="ca533-149">Type</span></span>

*   <span data-ttu-id="ca533-150">String</span><span class="sxs-lookup"><span data-stu-id="ca533-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca533-151">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-151">Requirements</span></span>

|<span data-ttu-id="ca533-152">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-152">Requirement</span></span>| <span data-ttu-id="ca533-153">值</span><span class="sxs-lookup"><span data-stu-id="ca533-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca533-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca533-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca533-155">1.0</span><span class="sxs-lookup"><span data-stu-id="ca533-155">1.0</span></span>|
|[<span data-ttu-id="ca533-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca533-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca533-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca533-157">ReadItem</span></span>|
|[<span data-ttu-id="ca533-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca533-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca533-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca533-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca533-160">示例</span><span class="sxs-lookup"><span data-stu-id="ca533-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ca533-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ca533-161">emailAddress :String</span></span>

<span data-ttu-id="ca533-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="ca533-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ca533-163">Type</span><span class="sxs-lookup"><span data-stu-id="ca533-163">Type</span></span>

*   <span data-ttu-id="ca533-164">String</span><span class="sxs-lookup"><span data-stu-id="ca533-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca533-165">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-165">Requirements</span></span>

|<span data-ttu-id="ca533-166">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-166">Requirement</span></span>| <span data-ttu-id="ca533-167">值</span><span class="sxs-lookup"><span data-stu-id="ca533-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca533-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca533-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca533-169">1.0</span><span class="sxs-lookup"><span data-stu-id="ca533-169">1.0</span></span>|
|[<span data-ttu-id="ca533-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca533-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca533-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca533-171">ReadItem</span></span>|
|[<span data-ttu-id="ca533-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca533-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca533-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca533-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca533-174">示例</span><span class="sxs-lookup"><span data-stu-id="ca533-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ca533-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ca533-175">timeZone :String</span></span>

<span data-ttu-id="ca533-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="ca533-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ca533-177">Type</span><span class="sxs-lookup"><span data-stu-id="ca533-177">Type</span></span>

*   <span data-ttu-id="ca533-178">String</span><span class="sxs-lookup"><span data-stu-id="ca533-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca533-179">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-179">Requirements</span></span>

|<span data-ttu-id="ca533-180">要求</span><span class="sxs-lookup"><span data-stu-id="ca533-180">Requirement</span></span>| <span data-ttu-id="ca533-181">值</span><span class="sxs-lookup"><span data-stu-id="ca533-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca533-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca533-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca533-183">1.0</span><span class="sxs-lookup"><span data-stu-id="ca533-183">1.0</span></span>|
|[<span data-ttu-id="ca533-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca533-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca533-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca533-185">ReadItem</span></span>|
|[<span data-ttu-id="ca533-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca533-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca533-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca533-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca533-188">示例</span><span class="sxs-lookup"><span data-stu-id="ca533-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
