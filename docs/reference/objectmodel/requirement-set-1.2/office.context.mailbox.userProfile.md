---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.2\""
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8ba2a21b16c51c827155d793241b80c5c510dd5a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696335"
---
# <a name="userprofile"></a><span data-ttu-id="29ef9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="29ef9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="29ef9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="29ef9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="29ef9-104">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-104">Requirements</span></span>

|<span data-ttu-id="29ef9-105">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-105">Requirement</span></span>| <span data-ttu-id="29ef9-106">值</span><span class="sxs-lookup"><span data-stu-id="29ef9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="29ef9-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="29ef9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29ef9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="29ef9-108">1.0</span></span>|
|[<span data-ttu-id="29ef9-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="29ef9-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29ef9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29ef9-110">ReadItem</span></span>|
|[<span data-ttu-id="29ef9-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="29ef9-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="29ef9-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="29ef9-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="29ef9-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="29ef9-113">Members and methods</span></span>

| <span data-ttu-id="29ef9-114">成员</span><span class="sxs-lookup"><span data-stu-id="29ef9-114">Member</span></span> | <span data-ttu-id="29ef9-115">类型</span><span class="sxs-lookup"><span data-stu-id="29ef9-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="29ef9-116">displayName</span><span class="sxs-lookup"><span data-stu-id="29ef9-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="29ef9-117">Member</span><span class="sxs-lookup"><span data-stu-id="29ef9-117">Member</span></span> |
| [<span data-ttu-id="29ef9-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="29ef9-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="29ef9-119">Member</span><span class="sxs-lookup"><span data-stu-id="29ef9-119">Member</span></span> |
| [<span data-ttu-id="29ef9-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="29ef9-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="29ef9-121">Member</span><span class="sxs-lookup"><span data-stu-id="29ef9-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="29ef9-122">Members</span><span class="sxs-lookup"><span data-stu-id="29ef9-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="29ef9-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="29ef9-123">displayName: String</span></span>

<span data-ttu-id="29ef9-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="29ef9-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="29ef9-125">类型</span><span class="sxs-lookup"><span data-stu-id="29ef9-125">Type</span></span>

*   <span data-ttu-id="29ef9-126">String</span><span class="sxs-lookup"><span data-stu-id="29ef9-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29ef9-127">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-127">Requirements</span></span>

|<span data-ttu-id="29ef9-128">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-128">Requirement</span></span>| <span data-ttu-id="29ef9-129">值</span><span class="sxs-lookup"><span data-stu-id="29ef9-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="29ef9-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="29ef9-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29ef9-131">1.0</span><span class="sxs-lookup"><span data-stu-id="29ef9-131">1.0</span></span>|
|[<span data-ttu-id="29ef9-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="29ef9-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29ef9-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29ef9-133">ReadItem</span></span>|
|[<span data-ttu-id="29ef9-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="29ef9-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="29ef9-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="29ef9-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29ef9-136">示例</span><span class="sxs-lookup"><span data-stu-id="29ef9-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="29ef9-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="29ef9-137">emailAddress: String</span></span>

<span data-ttu-id="29ef9-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="29ef9-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="29ef9-139">类型</span><span class="sxs-lookup"><span data-stu-id="29ef9-139">Type</span></span>

*   <span data-ttu-id="29ef9-140">String</span><span class="sxs-lookup"><span data-stu-id="29ef9-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29ef9-141">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-141">Requirements</span></span>

|<span data-ttu-id="29ef9-142">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-142">Requirement</span></span>| <span data-ttu-id="29ef9-143">值</span><span class="sxs-lookup"><span data-stu-id="29ef9-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="29ef9-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="29ef9-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29ef9-145">1.0</span><span class="sxs-lookup"><span data-stu-id="29ef9-145">1.0</span></span>|
|[<span data-ttu-id="29ef9-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="29ef9-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29ef9-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29ef9-147">ReadItem</span></span>|
|[<span data-ttu-id="29ef9-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="29ef9-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="29ef9-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="29ef9-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29ef9-150">示例</span><span class="sxs-lookup"><span data-stu-id="29ef9-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="29ef9-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="29ef9-151">timeZone: String</span></span>

<span data-ttu-id="29ef9-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="29ef9-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="29ef9-153">类型</span><span class="sxs-lookup"><span data-stu-id="29ef9-153">Type</span></span>

*   <span data-ttu-id="29ef9-154">String</span><span class="sxs-lookup"><span data-stu-id="29ef9-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29ef9-155">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-155">Requirements</span></span>

|<span data-ttu-id="29ef9-156">要求</span><span class="sxs-lookup"><span data-stu-id="29ef9-156">Requirement</span></span>| <span data-ttu-id="29ef9-157">值</span><span class="sxs-lookup"><span data-stu-id="29ef9-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="29ef9-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="29ef9-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29ef9-159">1.0</span><span class="sxs-lookup"><span data-stu-id="29ef9-159">1.0</span></span>|
|[<span data-ttu-id="29ef9-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="29ef9-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29ef9-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29ef9-161">ReadItem</span></span>|
|[<span data-ttu-id="29ef9-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="29ef9-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="29ef9-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="29ef9-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29ef9-164">示例</span><span class="sxs-lookup"><span data-stu-id="29ef9-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
