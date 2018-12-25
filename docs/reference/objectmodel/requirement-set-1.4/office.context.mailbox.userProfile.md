---
title: Office.context.mailbox.userProfile - 要求集 1.4
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 55d0a789c8e46fd3f6ee69f39cf33f7e7d94c322
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432646"
---
# <a name="userprofile"></a><span data-ttu-id="5c431-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="5c431-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="5c431-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="5c431-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c431-104">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-104">Requirements</span></span>

|<span data-ttu-id="5c431-105">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-105">Requirement</span></span>| <span data-ttu-id="5c431-106">值</span><span class="sxs-lookup"><span data-stu-id="5c431-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c431-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c431-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c431-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5c431-108">1.0</span></span>|
|[<span data-ttu-id="5c431-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c431-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c431-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c431-110">ReadItem</span></span>|
|[<span data-ttu-id="5c431-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c431-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c431-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c431-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="5c431-113">成员</span><span class="sxs-lookup"><span data-stu-id="5c431-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="5c431-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5c431-114">displayName :String</span></span>

<span data-ttu-id="5c431-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="5c431-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5c431-116">类型：</span><span class="sxs-lookup"><span data-stu-id="5c431-116">Type:</span></span>

*   <span data-ttu-id="5c431-117">String</span><span class="sxs-lookup"><span data-stu-id="5c431-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c431-118">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-118">Requirements</span></span>

|<span data-ttu-id="5c431-119">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-119">Requirement</span></span>| <span data-ttu-id="5c431-120">值</span><span class="sxs-lookup"><span data-stu-id="5c431-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c431-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c431-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c431-122">1.0</span><span class="sxs-lookup"><span data-stu-id="5c431-122">1.0</span></span>|
|[<span data-ttu-id="5c431-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c431-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c431-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c431-124">ReadItem</span></span>|
|[<span data-ttu-id="5c431-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c431-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c431-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c431-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c431-127">示例</span><span class="sxs-lookup"><span data-stu-id="5c431-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5c431-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5c431-128">emailAddress :String</span></span>

<span data-ttu-id="5c431-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="5c431-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5c431-130">类型：</span><span class="sxs-lookup"><span data-stu-id="5c431-130">Type:</span></span>

*   <span data-ttu-id="5c431-131">String</span><span class="sxs-lookup"><span data-stu-id="5c431-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c431-132">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-132">Requirements</span></span>

|<span data-ttu-id="5c431-133">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-133">Requirement</span></span>| <span data-ttu-id="5c431-134">值</span><span class="sxs-lookup"><span data-stu-id="5c431-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c431-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c431-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c431-136">1.0</span><span class="sxs-lookup"><span data-stu-id="5c431-136">1.0</span></span>|
|[<span data-ttu-id="5c431-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c431-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c431-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c431-138">ReadItem</span></span>|
|[<span data-ttu-id="5c431-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c431-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c431-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c431-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c431-141">示例</span><span class="sxs-lookup"><span data-stu-id="5c431-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5c431-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5c431-142">timeZone :String</span></span>

<span data-ttu-id="5c431-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="5c431-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5c431-144">类型：</span><span class="sxs-lookup"><span data-stu-id="5c431-144">Type:</span></span>

*   <span data-ttu-id="5c431-145">String</span><span class="sxs-lookup"><span data-stu-id="5c431-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c431-146">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-146">Requirements</span></span>

|<span data-ttu-id="5c431-147">要求</span><span class="sxs-lookup"><span data-stu-id="5c431-147">Requirement</span></span>| <span data-ttu-id="5c431-148">值</span><span class="sxs-lookup"><span data-stu-id="5c431-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c431-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5c431-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5c431-150">1.0</span><span class="sxs-lookup"><span data-stu-id="5c431-150">1.0</span></span>|
|[<span data-ttu-id="5c431-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5c431-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c431-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c431-152">ReadItem</span></span>|
|[<span data-ttu-id="5c431-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5c431-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5c431-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5c431-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5c431-155">示例</span><span class="sxs-lookup"><span data-stu-id="5c431-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```