---
title: Office.context.mailbox.userProfile - 要求集 1.1
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 312cba4d5aace980b7c9b205899fac51d3da3de5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433171"
---
# <a name="userprofile"></a><span data-ttu-id="c4b43-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c4b43-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c4b43-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c4b43-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4b43-104">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-104">Requirements</span></span>

|<span data-ttu-id="c4b43-105">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-105">Requirement</span></span>| <span data-ttu-id="c4b43-106">值</span><span class="sxs-lookup"><span data-stu-id="c4b43-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4b43-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c4b43-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4b43-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c4b43-108">1.0</span></span>|
|[<span data-ttu-id="c4b43-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4b43-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4b43-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4b43-110">ReadItem</span></span>|
|[<span data-ttu-id="c4b43-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4b43-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4b43-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4b43-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="c4b43-113">成员</span><span class="sxs-lookup"><span data-stu-id="c4b43-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c4b43-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c4b43-114">displayName :String</span></span>

<span data-ttu-id="c4b43-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="c4b43-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c4b43-116">类型：</span><span class="sxs-lookup"><span data-stu-id="c4b43-116">Type:</span></span>

*   <span data-ttu-id="c4b43-117">String</span><span class="sxs-lookup"><span data-stu-id="c4b43-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4b43-118">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-118">Requirements</span></span>

|<span data-ttu-id="c4b43-119">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-119">Requirement</span></span>| <span data-ttu-id="c4b43-120">值</span><span class="sxs-lookup"><span data-stu-id="c4b43-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4b43-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c4b43-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4b43-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c4b43-122">1.0</span></span>|
|[<span data-ttu-id="c4b43-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4b43-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4b43-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4b43-124">ReadItem</span></span>|
|[<span data-ttu-id="c4b43-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4b43-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4b43-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4b43-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4b43-127">示例</span><span class="sxs-lookup"><span data-stu-id="c4b43-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c4b43-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c4b43-128">emailAddress :String</span></span>

<span data-ttu-id="c4b43-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c4b43-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c4b43-130">类型：</span><span class="sxs-lookup"><span data-stu-id="c4b43-130">Type:</span></span>

*   <span data-ttu-id="c4b43-131">String</span><span class="sxs-lookup"><span data-stu-id="c4b43-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4b43-132">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-132">Requirements</span></span>

|<span data-ttu-id="c4b43-133">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-133">Requirement</span></span>| <span data-ttu-id="c4b43-134">值</span><span class="sxs-lookup"><span data-stu-id="c4b43-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4b43-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c4b43-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4b43-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c4b43-136">1.0</span></span>|
|[<span data-ttu-id="c4b43-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4b43-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4b43-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4b43-138">ReadItem</span></span>|
|[<span data-ttu-id="c4b43-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4b43-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4b43-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4b43-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4b43-141">示例</span><span class="sxs-lookup"><span data-stu-id="c4b43-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c4b43-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c4b43-142">timeZone :String</span></span>

<span data-ttu-id="c4b43-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="c4b43-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c4b43-144">类型：</span><span class="sxs-lookup"><span data-stu-id="c4b43-144">Type:</span></span>

*   <span data-ttu-id="c4b43-145">String</span><span class="sxs-lookup"><span data-stu-id="c4b43-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4b43-146">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-146">Requirements</span></span>

|<span data-ttu-id="c4b43-147">要求</span><span class="sxs-lookup"><span data-stu-id="c4b43-147">Requirement</span></span>| <span data-ttu-id="c4b43-148">值</span><span class="sxs-lookup"><span data-stu-id="c4b43-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4b43-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c4b43-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4b43-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c4b43-150">1.0</span></span>|
|[<span data-ttu-id="c4b43-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c4b43-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4b43-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4b43-152">ReadItem</span></span>|
|[<span data-ttu-id="c4b43-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c4b43-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4b43-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c4b43-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4b43-155">示例</span><span class="sxs-lookup"><span data-stu-id="c4b43-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```