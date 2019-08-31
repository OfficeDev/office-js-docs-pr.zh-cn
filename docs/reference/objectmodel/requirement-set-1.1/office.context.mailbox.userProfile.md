---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.1\""
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 06492623e0b9ab16792d6b23dfaeb27d99125ff1
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696398"
---
# <a name="userprofile"></a><span data-ttu-id="e2c95-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e2c95-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e2c95-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e2c95-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2c95-104">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-104">Requirements</span></span>

|<span data-ttu-id="e2c95-105">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-105">Requirement</span></span>| <span data-ttu-id="e2c95-106">值</span><span class="sxs-lookup"><span data-stu-id="e2c95-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2c95-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e2c95-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2c95-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e2c95-108">1.0</span></span>|
|[<span data-ttu-id="e2c95-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e2c95-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2c95-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2c95-110">ReadItem</span></span>|
|[<span data-ttu-id="e2c95-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e2c95-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2c95-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e2c95-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e2c95-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="e2c95-113">Members and methods</span></span>

| <span data-ttu-id="e2c95-114">成员</span><span class="sxs-lookup"><span data-stu-id="e2c95-114">Member</span></span> | <span data-ttu-id="e2c95-115">类型</span><span class="sxs-lookup"><span data-stu-id="e2c95-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e2c95-116">displayName</span><span class="sxs-lookup"><span data-stu-id="e2c95-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="e2c95-117">Member</span><span class="sxs-lookup"><span data-stu-id="e2c95-117">Member</span></span> |
| [<span data-ttu-id="e2c95-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e2c95-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e2c95-119">Member</span><span class="sxs-lookup"><span data-stu-id="e2c95-119">Member</span></span> |
| [<span data-ttu-id="e2c95-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="e2c95-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e2c95-121">Member</span><span class="sxs-lookup"><span data-stu-id="e2c95-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e2c95-122">Members</span><span class="sxs-lookup"><span data-stu-id="e2c95-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="e2c95-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="e2c95-123">displayName: String</span></span>

<span data-ttu-id="e2c95-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="e2c95-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e2c95-125">类型</span><span class="sxs-lookup"><span data-stu-id="e2c95-125">Type</span></span>

*   <span data-ttu-id="e2c95-126">String</span><span class="sxs-lookup"><span data-stu-id="e2c95-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2c95-127">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-127">Requirements</span></span>

|<span data-ttu-id="e2c95-128">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-128">Requirement</span></span>| <span data-ttu-id="e2c95-129">值</span><span class="sxs-lookup"><span data-stu-id="e2c95-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2c95-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e2c95-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2c95-131">1.0</span><span class="sxs-lookup"><span data-stu-id="e2c95-131">1.0</span></span>|
|[<span data-ttu-id="e2c95-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e2c95-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2c95-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2c95-133">ReadItem</span></span>|
|[<span data-ttu-id="e2c95-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e2c95-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2c95-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e2c95-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2c95-136">示例</span><span class="sxs-lookup"><span data-stu-id="e2c95-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="e2c95-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="e2c95-137">emailAddress: String</span></span>

<span data-ttu-id="e2c95-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="e2c95-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e2c95-139">类型</span><span class="sxs-lookup"><span data-stu-id="e2c95-139">Type</span></span>

*   <span data-ttu-id="e2c95-140">String</span><span class="sxs-lookup"><span data-stu-id="e2c95-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2c95-141">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-141">Requirements</span></span>

|<span data-ttu-id="e2c95-142">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-142">Requirement</span></span>| <span data-ttu-id="e2c95-143">值</span><span class="sxs-lookup"><span data-stu-id="e2c95-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2c95-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e2c95-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2c95-145">1.0</span><span class="sxs-lookup"><span data-stu-id="e2c95-145">1.0</span></span>|
|[<span data-ttu-id="e2c95-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e2c95-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2c95-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2c95-147">ReadItem</span></span>|
|[<span data-ttu-id="e2c95-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e2c95-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2c95-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e2c95-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2c95-150">示例</span><span class="sxs-lookup"><span data-stu-id="e2c95-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="e2c95-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="e2c95-151">timeZone: String</span></span>

<span data-ttu-id="e2c95-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="e2c95-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e2c95-153">类型</span><span class="sxs-lookup"><span data-stu-id="e2c95-153">Type</span></span>

*   <span data-ttu-id="e2c95-154">String</span><span class="sxs-lookup"><span data-stu-id="e2c95-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2c95-155">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-155">Requirements</span></span>

|<span data-ttu-id="e2c95-156">要求</span><span class="sxs-lookup"><span data-stu-id="e2c95-156">Requirement</span></span>| <span data-ttu-id="e2c95-157">值</span><span class="sxs-lookup"><span data-stu-id="e2c95-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2c95-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e2c95-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2c95-159">1.0</span><span class="sxs-lookup"><span data-stu-id="e2c95-159">1.0</span></span>|
|[<span data-ttu-id="e2c95-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e2c95-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2c95-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2c95-161">ReadItem</span></span>|
|[<span data-ttu-id="e2c95-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e2c95-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2c95-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e2c95-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2c95-164">示例</span><span class="sxs-lookup"><span data-stu-id="e2c95-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
