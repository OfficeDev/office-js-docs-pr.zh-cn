---
title: Office.context.mailbox.userProfile - 要求集 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/13/2019
ms.locfileid: "29635900"
---
# <a name="userprofile"></a><span data-ttu-id="cb1a7-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="cb1a7-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="cb1a7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="cb1a7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1a7-104">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-104">Requirements</span></span>

|<span data-ttu-id="cb1a7-105">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-105">Requirement</span></span>| <span data-ttu-id="cb1a7-106">值</span><span class="sxs-lookup"><span data-stu-id="cb1a7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1a7-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cb1a7-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1a7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="cb1a7-108">1.0</span></span>|
|[<span data-ttu-id="cb1a7-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cb1a7-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1a7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1a7-110">ReadItem</span></span>|
|[<span data-ttu-id="cb1a7-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cb1a7-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1a7-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cb1a7-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cb1a7-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="cb1a7-113">Members and methods</span></span>

| <span data-ttu-id="cb1a7-114">成员</span><span class="sxs-lookup"><span data-stu-id="cb1a7-114">Member</span></span> | <span data-ttu-id="cb1a7-115">类型</span><span class="sxs-lookup"><span data-stu-id="cb1a7-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cb1a7-116">accountType</span><span class="sxs-lookup"><span data-stu-id="cb1a7-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="cb1a7-117">成员</span><span class="sxs-lookup"><span data-stu-id="cb1a7-117">Member</span></span> |
| [<span data-ttu-id="cb1a7-118">displayName</span><span class="sxs-lookup"><span data-stu-id="cb1a7-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="cb1a7-119">Member</span><span class="sxs-lookup"><span data-stu-id="cb1a7-119">Member</span></span> |
| [<span data-ttu-id="cb1a7-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="cb1a7-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="cb1a7-121">Member</span><span class="sxs-lookup"><span data-stu-id="cb1a7-121">Member</span></span> |
| [<span data-ttu-id="cb1a7-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="cb1a7-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="cb1a7-123">Member</span><span class="sxs-lookup"><span data-stu-id="cb1a7-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="cb1a7-124">成员</span><span class="sxs-lookup"><span data-stu-id="cb1a7-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="cb1a7-125">accountType：字符串</span><span class="sxs-lookup"><span data-stu-id="cb1a7-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="cb1a7-126">此成员是当前只支持 for Mac Outlook 2016 (生成 16.9.1212 或更高版本)。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="cb1a7-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="cb1a7-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="cb1a7-129">值</span><span class="sxs-lookup"><span data-stu-id="cb1a7-129">Value</span></span> | <span data-ttu-id="cb1a7-130">说明</span><span class="sxs-lookup"><span data-stu-id="cb1a7-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="cb1a7-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="cb1a7-132">邮箱与 Gmail 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="cb1a7-133">邮箱与 Office 365 工作或学校帐户关联。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="cb1a7-134">邮箱与个人 Outlook.com 帐户关联。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="cb1a7-135">类型：</span><span class="sxs-lookup"><span data-stu-id="cb1a7-135">Type:</span></span>

*   <span data-ttu-id="cb1a7-136">String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1a7-137">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-137">Requirements</span></span>

|<span data-ttu-id="cb1a7-138">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-138">Requirement</span></span>| <span data-ttu-id="cb1a7-139">值</span><span class="sxs-lookup"><span data-stu-id="cb1a7-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1a7-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cb1a7-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1a7-141">1.6</span><span class="sxs-lookup"><span data-stu-id="cb1a7-141">1.6</span></span> |
|[<span data-ttu-id="cb1a7-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cb1a7-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1a7-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1a7-143">ReadItem</span></span>|
|[<span data-ttu-id="cb1a7-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cb1a7-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1a7-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cb1a7-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1a7-146">示例</span><span class="sxs-lookup"><span data-stu-id="cb1a7-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="cb1a7-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-147">displayName :String</span></span>

<span data-ttu-id="cb1a7-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1a7-149">类型：</span><span class="sxs-lookup"><span data-stu-id="cb1a7-149">Type:</span></span>

*   <span data-ttu-id="cb1a7-150">String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1a7-151">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-151">Requirements</span></span>

|<span data-ttu-id="cb1a7-152">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-152">Requirement</span></span>| <span data-ttu-id="cb1a7-153">值</span><span class="sxs-lookup"><span data-stu-id="cb1a7-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1a7-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cb1a7-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1a7-155">1.0</span><span class="sxs-lookup"><span data-stu-id="cb1a7-155">1.0</span></span>|
|[<span data-ttu-id="cb1a7-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cb1a7-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1a7-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1a7-157">ReadItem</span></span>|
|[<span data-ttu-id="cb1a7-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cb1a7-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1a7-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cb1a7-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1a7-160">示例</span><span class="sxs-lookup"><span data-stu-id="cb1a7-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="cb1a7-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-161">emailAddress :String</span></span>

<span data-ttu-id="cb1a7-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1a7-163">类型：</span><span class="sxs-lookup"><span data-stu-id="cb1a7-163">Type:</span></span>

*   <span data-ttu-id="cb1a7-164">String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1a7-165">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-165">Requirements</span></span>

|<span data-ttu-id="cb1a7-166">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-166">Requirement</span></span>| <span data-ttu-id="cb1a7-167">值</span><span class="sxs-lookup"><span data-stu-id="cb1a7-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1a7-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cb1a7-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1a7-169">1.0</span><span class="sxs-lookup"><span data-stu-id="cb1a7-169">1.0</span></span>|
|[<span data-ttu-id="cb1a7-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cb1a7-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1a7-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1a7-171">ReadItem</span></span>|
|[<span data-ttu-id="cb1a7-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cb1a7-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1a7-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cb1a7-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1a7-174">示例</span><span class="sxs-lookup"><span data-stu-id="cb1a7-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="cb1a7-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-175">timeZone :String</span></span>

<span data-ttu-id="cb1a7-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="cb1a7-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1a7-177">类型：</span><span class="sxs-lookup"><span data-stu-id="cb1a7-177">Type:</span></span>

*   <span data-ttu-id="cb1a7-178">String</span><span class="sxs-lookup"><span data-stu-id="cb1a7-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1a7-179">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-179">Requirements</span></span>

|<span data-ttu-id="cb1a7-180">要求</span><span class="sxs-lookup"><span data-stu-id="cb1a7-180">Requirement</span></span>| <span data-ttu-id="cb1a7-181">值</span><span class="sxs-lookup"><span data-stu-id="cb1a7-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1a7-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cb1a7-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1a7-183">1.0</span><span class="sxs-lookup"><span data-stu-id="cb1a7-183">1.0</span></span>|
|[<span data-ttu-id="cb1a7-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cb1a7-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1a7-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1a7-185">ReadItem</span></span>|
|[<span data-ttu-id="cb1a7-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cb1a7-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1a7-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cb1a7-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1a7-188">示例</span><span class="sxs-lookup"><span data-stu-id="cb1a7-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
