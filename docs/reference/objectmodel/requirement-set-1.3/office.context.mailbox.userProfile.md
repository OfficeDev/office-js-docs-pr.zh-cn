---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.3\""
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 20393d0ac650de34054b912d9e53a9ac167fddb2
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696251"
---
# <a name="userprofile"></a><span data-ttu-id="0f2ac-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="0f2ac-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="0f2ac-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="0f2ac-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f2ac-104">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-104">Requirements</span></span>

|<span data-ttu-id="0f2ac-105">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-105">Requirement</span></span>| <span data-ttu-id="0f2ac-106">值</span><span class="sxs-lookup"><span data-stu-id="0f2ac-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f2ac-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0f2ac-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f2ac-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0f2ac-108">1.0</span></span>|
|[<span data-ttu-id="0f2ac-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0f2ac-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f2ac-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f2ac-110">ReadItem</span></span>|
|[<span data-ttu-id="0f2ac-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0f2ac-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f2ac-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0f2ac-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0f2ac-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="0f2ac-113">Members and methods</span></span>

| <span data-ttu-id="0f2ac-114">成员</span><span class="sxs-lookup"><span data-stu-id="0f2ac-114">Member</span></span> | <span data-ttu-id="0f2ac-115">类型</span><span class="sxs-lookup"><span data-stu-id="0f2ac-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0f2ac-116">displayName</span><span class="sxs-lookup"><span data-stu-id="0f2ac-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="0f2ac-117">Member</span><span class="sxs-lookup"><span data-stu-id="0f2ac-117">Member</span></span> |
| [<span data-ttu-id="0f2ac-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="0f2ac-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="0f2ac-119">Member</span><span class="sxs-lookup"><span data-stu-id="0f2ac-119">Member</span></span> |
| [<span data-ttu-id="0f2ac-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="0f2ac-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="0f2ac-121">Member</span><span class="sxs-lookup"><span data-stu-id="0f2ac-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="0f2ac-122">Members</span><span class="sxs-lookup"><span data-stu-id="0f2ac-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="0f2ac-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="0f2ac-123">displayName: String</span></span>

<span data-ttu-id="0f2ac-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="0f2ac-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="0f2ac-125">类型</span><span class="sxs-lookup"><span data-stu-id="0f2ac-125">Type</span></span>

*   <span data-ttu-id="0f2ac-126">String</span><span class="sxs-lookup"><span data-stu-id="0f2ac-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f2ac-127">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-127">Requirements</span></span>

|<span data-ttu-id="0f2ac-128">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-128">Requirement</span></span>| <span data-ttu-id="0f2ac-129">值</span><span class="sxs-lookup"><span data-stu-id="0f2ac-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f2ac-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0f2ac-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f2ac-131">1.0</span><span class="sxs-lookup"><span data-stu-id="0f2ac-131">1.0</span></span>|
|[<span data-ttu-id="0f2ac-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0f2ac-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f2ac-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f2ac-133">ReadItem</span></span>|
|[<span data-ttu-id="0f2ac-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0f2ac-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f2ac-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0f2ac-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f2ac-136">示例</span><span class="sxs-lookup"><span data-stu-id="0f2ac-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="0f2ac-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="0f2ac-137">emailAddress: String</span></span>

<span data-ttu-id="0f2ac-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="0f2ac-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="0f2ac-139">类型</span><span class="sxs-lookup"><span data-stu-id="0f2ac-139">Type</span></span>

*   <span data-ttu-id="0f2ac-140">String</span><span class="sxs-lookup"><span data-stu-id="0f2ac-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f2ac-141">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-141">Requirements</span></span>

|<span data-ttu-id="0f2ac-142">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-142">Requirement</span></span>| <span data-ttu-id="0f2ac-143">值</span><span class="sxs-lookup"><span data-stu-id="0f2ac-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f2ac-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0f2ac-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f2ac-145">1.0</span><span class="sxs-lookup"><span data-stu-id="0f2ac-145">1.0</span></span>|
|[<span data-ttu-id="0f2ac-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0f2ac-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f2ac-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f2ac-147">ReadItem</span></span>|
|[<span data-ttu-id="0f2ac-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0f2ac-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f2ac-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0f2ac-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f2ac-150">示例</span><span class="sxs-lookup"><span data-stu-id="0f2ac-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="0f2ac-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="0f2ac-151">timeZone: String</span></span>

<span data-ttu-id="0f2ac-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="0f2ac-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="0f2ac-153">类型</span><span class="sxs-lookup"><span data-stu-id="0f2ac-153">Type</span></span>

*   <span data-ttu-id="0f2ac-154">String</span><span class="sxs-lookup"><span data-stu-id="0f2ac-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f2ac-155">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-155">Requirements</span></span>

|<span data-ttu-id="0f2ac-156">要求</span><span class="sxs-lookup"><span data-stu-id="0f2ac-156">Requirement</span></span>| <span data-ttu-id="0f2ac-157">值</span><span class="sxs-lookup"><span data-stu-id="0f2ac-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f2ac-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0f2ac-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f2ac-159">1.0</span><span class="sxs-lookup"><span data-stu-id="0f2ac-159">1.0</span></span>|
|[<span data-ttu-id="0f2ac-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0f2ac-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f2ac-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f2ac-161">ReadItem</span></span>|
|[<span data-ttu-id="0f2ac-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0f2ac-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f2ac-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0f2ac-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f2ac-164">示例</span><span class="sxs-lookup"><span data-stu-id="0f2ac-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
