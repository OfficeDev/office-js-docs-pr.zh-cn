---
title: Office.context.mailbox.userProfile - 要求集 1.2
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: e5548fa514cff9b452c2747324f11e5df8a06def
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432240"
---
# <a name="userprofile"></a><span data-ttu-id="ca475-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ca475-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ca475-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ca475-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca475-104">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-104">Requirements</span></span>

|<span data-ttu-id="ca475-105">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-105">Requirement</span></span>| <span data-ttu-id="ca475-106">值</span><span class="sxs-lookup"><span data-stu-id="ca475-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca475-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca475-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca475-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ca475-108">1.0</span></span>|
|[<span data-ttu-id="ca475-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca475-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca475-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca475-110">ReadItem</span></span>|
|[<span data-ttu-id="ca475-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca475-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca475-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca475-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="ca475-113">成员</span><span class="sxs-lookup"><span data-stu-id="ca475-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ca475-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ca475-114">displayName :String</span></span>

<span data-ttu-id="ca475-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="ca475-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ca475-116">类型：</span><span class="sxs-lookup"><span data-stu-id="ca475-116">Type:</span></span>

*   <span data-ttu-id="ca475-117">String</span><span class="sxs-lookup"><span data-stu-id="ca475-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca475-118">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-118">Requirements</span></span>

|<span data-ttu-id="ca475-119">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-119">Requirement</span></span>| <span data-ttu-id="ca475-120">值</span><span class="sxs-lookup"><span data-stu-id="ca475-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca475-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca475-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca475-122">1.0</span><span class="sxs-lookup"><span data-stu-id="ca475-122">1.0</span></span>|
|[<span data-ttu-id="ca475-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca475-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca475-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca475-124">ReadItem</span></span>|
|[<span data-ttu-id="ca475-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca475-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca475-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca475-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca475-127">示例</span><span class="sxs-lookup"><span data-stu-id="ca475-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ca475-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ca475-128">emailAddress :String</span></span>

<span data-ttu-id="ca475-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="ca475-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ca475-130">类型：</span><span class="sxs-lookup"><span data-stu-id="ca475-130">Type:</span></span>

*   <span data-ttu-id="ca475-131">String</span><span class="sxs-lookup"><span data-stu-id="ca475-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca475-132">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-132">Requirements</span></span>

|<span data-ttu-id="ca475-133">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-133">Requirement</span></span>| <span data-ttu-id="ca475-134">值</span><span class="sxs-lookup"><span data-stu-id="ca475-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca475-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca475-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca475-136">1.0</span><span class="sxs-lookup"><span data-stu-id="ca475-136">1.0</span></span>|
|[<span data-ttu-id="ca475-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca475-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca475-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca475-138">ReadItem</span></span>|
|[<span data-ttu-id="ca475-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca475-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca475-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca475-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca475-141">示例</span><span class="sxs-lookup"><span data-stu-id="ca475-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ca475-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ca475-142">timeZone :String</span></span>

<span data-ttu-id="ca475-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="ca475-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ca475-144">类型：</span><span class="sxs-lookup"><span data-stu-id="ca475-144">Type:</span></span>

*   <span data-ttu-id="ca475-145">String</span><span class="sxs-lookup"><span data-stu-id="ca475-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca475-146">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-146">Requirements</span></span>

|<span data-ttu-id="ca475-147">要求</span><span class="sxs-lookup"><span data-stu-id="ca475-147">Requirement</span></span>| <span data-ttu-id="ca475-148">值</span><span class="sxs-lookup"><span data-stu-id="ca475-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca475-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca475-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca475-150">1.0</span><span class="sxs-lookup"><span data-stu-id="ca475-150">1.0</span></span>|
|[<span data-ttu-id="ca475-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca475-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca475-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca475-152">ReadItem</span></span>|
|[<span data-ttu-id="ca475-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca475-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca475-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca475-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca475-155">示例</span><span class="sxs-lookup"><span data-stu-id="ca475-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```