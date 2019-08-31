---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.5\""
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 993fad674fcc616483ac927619e7ca64d81b7326
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696090"
---
# <a name="userprofile"></a><span data-ttu-id="21630-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="21630-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="21630-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="21630-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="21630-104">要求</span><span class="sxs-lookup"><span data-stu-id="21630-104">Requirements</span></span>

|<span data-ttu-id="21630-105">要求</span><span class="sxs-lookup"><span data-stu-id="21630-105">Requirement</span></span>| <span data-ttu-id="21630-106">值</span><span class="sxs-lookup"><span data-stu-id="21630-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="21630-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="21630-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21630-108">1.0</span><span class="sxs-lookup"><span data-stu-id="21630-108">1.0</span></span>|
|[<span data-ttu-id="21630-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="21630-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21630-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21630-110">ReadItem</span></span>|
|[<span data-ttu-id="21630-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="21630-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="21630-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="21630-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="21630-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="21630-113">Members and methods</span></span>

| <span data-ttu-id="21630-114">成员</span><span class="sxs-lookup"><span data-stu-id="21630-114">Member</span></span> | <span data-ttu-id="21630-115">类型</span><span class="sxs-lookup"><span data-stu-id="21630-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="21630-116">displayName</span><span class="sxs-lookup"><span data-stu-id="21630-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="21630-117">Member</span><span class="sxs-lookup"><span data-stu-id="21630-117">Member</span></span> |
| [<span data-ttu-id="21630-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="21630-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="21630-119">Member</span><span class="sxs-lookup"><span data-stu-id="21630-119">Member</span></span> |
| [<span data-ttu-id="21630-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="21630-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="21630-121">Member</span><span class="sxs-lookup"><span data-stu-id="21630-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="21630-122">Members</span><span class="sxs-lookup"><span data-stu-id="21630-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="21630-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="21630-123">displayName: String</span></span>

<span data-ttu-id="21630-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="21630-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="21630-125">类型</span><span class="sxs-lookup"><span data-stu-id="21630-125">Type</span></span>

*   <span data-ttu-id="21630-126">String</span><span class="sxs-lookup"><span data-stu-id="21630-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21630-127">要求</span><span class="sxs-lookup"><span data-stu-id="21630-127">Requirements</span></span>

|<span data-ttu-id="21630-128">要求</span><span class="sxs-lookup"><span data-stu-id="21630-128">Requirement</span></span>| <span data-ttu-id="21630-129">值</span><span class="sxs-lookup"><span data-stu-id="21630-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="21630-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="21630-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21630-131">1.0</span><span class="sxs-lookup"><span data-stu-id="21630-131">1.0</span></span>|
|[<span data-ttu-id="21630-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="21630-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21630-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21630-133">ReadItem</span></span>|
|[<span data-ttu-id="21630-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="21630-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="21630-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="21630-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21630-136">示例</span><span class="sxs-lookup"><span data-stu-id="21630-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="21630-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="21630-137">emailAddress: String</span></span>

<span data-ttu-id="21630-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="21630-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="21630-139">类型</span><span class="sxs-lookup"><span data-stu-id="21630-139">Type</span></span>

*   <span data-ttu-id="21630-140">String</span><span class="sxs-lookup"><span data-stu-id="21630-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21630-141">要求</span><span class="sxs-lookup"><span data-stu-id="21630-141">Requirements</span></span>

|<span data-ttu-id="21630-142">要求</span><span class="sxs-lookup"><span data-stu-id="21630-142">Requirement</span></span>| <span data-ttu-id="21630-143">值</span><span class="sxs-lookup"><span data-stu-id="21630-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="21630-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="21630-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21630-145">1.0</span><span class="sxs-lookup"><span data-stu-id="21630-145">1.0</span></span>|
|[<span data-ttu-id="21630-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="21630-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21630-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21630-147">ReadItem</span></span>|
|[<span data-ttu-id="21630-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="21630-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="21630-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="21630-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21630-150">示例</span><span class="sxs-lookup"><span data-stu-id="21630-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="21630-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="21630-151">timeZone: String</span></span>

<span data-ttu-id="21630-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="21630-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="21630-153">类型</span><span class="sxs-lookup"><span data-stu-id="21630-153">Type</span></span>

*   <span data-ttu-id="21630-154">String</span><span class="sxs-lookup"><span data-stu-id="21630-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21630-155">要求</span><span class="sxs-lookup"><span data-stu-id="21630-155">Requirements</span></span>

|<span data-ttu-id="21630-156">要求</span><span class="sxs-lookup"><span data-stu-id="21630-156">Requirement</span></span>| <span data-ttu-id="21630-157">值</span><span class="sxs-lookup"><span data-stu-id="21630-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="21630-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="21630-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21630-159">1.0</span><span class="sxs-lookup"><span data-stu-id="21630-159">1.0</span></span>|
|[<span data-ttu-id="21630-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="21630-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21630-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21630-161">ReadItem</span></span>|
|[<span data-ttu-id="21630-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="21630-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="21630-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="21630-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="21630-164">示例</span><span class="sxs-lookup"><span data-stu-id="21630-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
