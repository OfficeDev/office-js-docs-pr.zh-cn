---
title: Office.context.mailbox.userProfile - 要求集 1.3
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 9f36b5f1d31ad6709cf2c43ce7dcb3f91a35bd00
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432219"
---
# <a name="userprofile"></a><span data-ttu-id="572e2-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="572e2-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="572e2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="572e2-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="572e2-104">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-104">Requirements</span></span>

|<span data-ttu-id="572e2-105">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-105">Requirement</span></span>| <span data-ttu-id="572e2-106">值</span><span class="sxs-lookup"><span data-stu-id="572e2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="572e2-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="572e2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="572e2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="572e2-108">1.0</span></span>|
|[<span data-ttu-id="572e2-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="572e2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="572e2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="572e2-110">ReadItem</span></span>|
|[<span data-ttu-id="572e2-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="572e2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="572e2-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="572e2-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="572e2-113">成员</span><span class="sxs-lookup"><span data-stu-id="572e2-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="572e2-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="572e2-114">displayName :String</span></span>

<span data-ttu-id="572e2-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="572e2-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="572e2-116">类型：</span><span class="sxs-lookup"><span data-stu-id="572e2-116">Type:</span></span>

*   <span data-ttu-id="572e2-117">String</span><span class="sxs-lookup"><span data-stu-id="572e2-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="572e2-118">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-118">Requirements</span></span>

|<span data-ttu-id="572e2-119">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-119">Requirement</span></span>| <span data-ttu-id="572e2-120">值</span><span class="sxs-lookup"><span data-stu-id="572e2-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="572e2-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="572e2-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="572e2-122">1.0</span><span class="sxs-lookup"><span data-stu-id="572e2-122">1.0</span></span>|
|[<span data-ttu-id="572e2-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="572e2-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="572e2-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="572e2-124">ReadItem</span></span>|
|[<span data-ttu-id="572e2-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="572e2-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="572e2-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="572e2-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="572e2-127">示例</span><span class="sxs-lookup"><span data-stu-id="572e2-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="572e2-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="572e2-128">emailAddress :String</span></span>

<span data-ttu-id="572e2-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="572e2-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="572e2-130">类型：</span><span class="sxs-lookup"><span data-stu-id="572e2-130">Type:</span></span>

*   <span data-ttu-id="572e2-131">String</span><span class="sxs-lookup"><span data-stu-id="572e2-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="572e2-132">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-132">Requirements</span></span>

|<span data-ttu-id="572e2-133">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-133">Requirement</span></span>| <span data-ttu-id="572e2-134">值</span><span class="sxs-lookup"><span data-stu-id="572e2-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="572e2-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="572e2-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="572e2-136">1.0</span><span class="sxs-lookup"><span data-stu-id="572e2-136">1.0</span></span>|
|[<span data-ttu-id="572e2-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="572e2-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="572e2-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="572e2-138">ReadItem</span></span>|
|[<span data-ttu-id="572e2-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="572e2-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="572e2-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="572e2-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="572e2-141">示例</span><span class="sxs-lookup"><span data-stu-id="572e2-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="572e2-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="572e2-142">timeZone :String</span></span>

<span data-ttu-id="572e2-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="572e2-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="572e2-144">类型：</span><span class="sxs-lookup"><span data-stu-id="572e2-144">Type:</span></span>

*   <span data-ttu-id="572e2-145">String</span><span class="sxs-lookup"><span data-stu-id="572e2-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="572e2-146">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-146">Requirements</span></span>

|<span data-ttu-id="572e2-147">要求</span><span class="sxs-lookup"><span data-stu-id="572e2-147">Requirement</span></span>| <span data-ttu-id="572e2-148">值</span><span class="sxs-lookup"><span data-stu-id="572e2-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="572e2-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="572e2-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="572e2-150">1.0</span><span class="sxs-lookup"><span data-stu-id="572e2-150">1.0</span></span>|
|[<span data-ttu-id="572e2-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="572e2-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="572e2-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="572e2-152">ReadItem</span></span>|
|[<span data-ttu-id="572e2-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="572e2-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="572e2-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="572e2-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="572e2-155">示例</span><span class="sxs-lookup"><span data-stu-id="572e2-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```