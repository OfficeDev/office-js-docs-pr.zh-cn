---
title: Office.context.mailbox.userProfile - 要求集 1.5
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 748daf4d14aae1d14560d29e1d76eeea09830573
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432716"
---
# <a name="userprofile"></a><span data-ttu-id="484ad-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="484ad-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="484ad-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="484ad-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="484ad-104">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-104">Requirements</span></span>

|<span data-ttu-id="484ad-105">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-105">Requirement</span></span>| <span data-ttu-id="484ad-106">值</span><span class="sxs-lookup"><span data-stu-id="484ad-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="484ad-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="484ad-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484ad-108">1.0</span><span class="sxs-lookup"><span data-stu-id="484ad-108">1.0</span></span>|
|[<span data-ttu-id="484ad-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="484ad-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484ad-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484ad-110">ReadItem</span></span>|
|[<span data-ttu-id="484ad-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="484ad-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484ad-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="484ad-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="484ad-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="484ad-113">Members and methods</span></span>

| <span data-ttu-id="484ad-114">成员</span><span class="sxs-lookup"><span data-stu-id="484ad-114">Member</span></span> | <span data-ttu-id="484ad-115">类型</span><span class="sxs-lookup"><span data-stu-id="484ad-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="484ad-116">displayName</span><span class="sxs-lookup"><span data-stu-id="484ad-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="484ad-117">成员</span><span class="sxs-lookup"><span data-stu-id="484ad-117">Member</span></span> |
| [<span data-ttu-id="484ad-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="484ad-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="484ad-119">成员</span><span class="sxs-lookup"><span data-stu-id="484ad-119">Member</span></span> |
| [<span data-ttu-id="484ad-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="484ad-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="484ad-121">成员</span><span class="sxs-lookup"><span data-stu-id="484ad-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="484ad-122">成员</span><span class="sxs-lookup"><span data-stu-id="484ad-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="484ad-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="484ad-123">displayName :String</span></span>

<span data-ttu-id="484ad-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="484ad-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="484ad-125">类型：</span><span class="sxs-lookup"><span data-stu-id="484ad-125">Type:</span></span>

*   <span data-ttu-id="484ad-126">String</span><span class="sxs-lookup"><span data-stu-id="484ad-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484ad-127">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-127">Requirements</span></span>

|<span data-ttu-id="484ad-128">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-128">Requirement</span></span>| <span data-ttu-id="484ad-129">值</span><span class="sxs-lookup"><span data-stu-id="484ad-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="484ad-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="484ad-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484ad-131">1.0</span><span class="sxs-lookup"><span data-stu-id="484ad-131">1.0</span></span>|
|[<span data-ttu-id="484ad-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="484ad-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484ad-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484ad-133">ReadItem</span></span>|
|[<span data-ttu-id="484ad-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="484ad-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484ad-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="484ad-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484ad-136">示例</span><span class="sxs-lookup"><span data-stu-id="484ad-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="484ad-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="484ad-137">emailAddress :String</span></span>

<span data-ttu-id="484ad-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="484ad-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="484ad-139">类型：</span><span class="sxs-lookup"><span data-stu-id="484ad-139">Type:</span></span>

*   <span data-ttu-id="484ad-140">String</span><span class="sxs-lookup"><span data-stu-id="484ad-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484ad-141">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-141">Requirements</span></span>

|<span data-ttu-id="484ad-142">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-142">Requirement</span></span>| <span data-ttu-id="484ad-143">值</span><span class="sxs-lookup"><span data-stu-id="484ad-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="484ad-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="484ad-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484ad-145">1.0</span><span class="sxs-lookup"><span data-stu-id="484ad-145">1.0</span></span>|
|[<span data-ttu-id="484ad-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="484ad-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484ad-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484ad-147">ReadItem</span></span>|
|[<span data-ttu-id="484ad-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="484ad-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484ad-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="484ad-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484ad-150">示例</span><span class="sxs-lookup"><span data-stu-id="484ad-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="484ad-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="484ad-151">timeZone :String</span></span>

<span data-ttu-id="484ad-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="484ad-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="484ad-153">类型：</span><span class="sxs-lookup"><span data-stu-id="484ad-153">Type:</span></span>

*   <span data-ttu-id="484ad-154">String</span><span class="sxs-lookup"><span data-stu-id="484ad-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484ad-155">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-155">Requirements</span></span>

|<span data-ttu-id="484ad-156">要求</span><span class="sxs-lookup"><span data-stu-id="484ad-156">Requirement</span></span>| <span data-ttu-id="484ad-157">值</span><span class="sxs-lookup"><span data-stu-id="484ad-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="484ad-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="484ad-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484ad-159">1.0</span><span class="sxs-lookup"><span data-stu-id="484ad-159">1.0</span></span>|
|[<span data-ttu-id="484ad-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="484ad-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484ad-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484ad-161">ReadItem</span></span>|
|[<span data-ttu-id="484ad-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="484ad-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484ad-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="484ad-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484ad-164">示例</span><span class="sxs-lookup"><span data-stu-id="484ad-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```