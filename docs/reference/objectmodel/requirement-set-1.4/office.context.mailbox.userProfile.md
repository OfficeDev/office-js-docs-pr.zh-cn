---
title: Office.context.mailbox.userProfile - 要求集 1.4
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 7facc0ea555dca7d6784a09f798c3d8fa25f2731
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067844"
---
# <a name="userprofile"></a><span data-ttu-id="ff7d2-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ff7d2-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ff7d2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ff7d2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff7d2-104">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-104">Requirements</span></span>

|<span data-ttu-id="ff7d2-105">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-105">Requirement</span></span>| <span data-ttu-id="ff7d2-106">值</span><span class="sxs-lookup"><span data-stu-id="ff7d2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff7d2-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ff7d2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff7d2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ff7d2-108">1.0</span></span>|
|[<span data-ttu-id="ff7d2-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ff7d2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff7d2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff7d2-110">ReadItem</span></span>|
|[<span data-ttu-id="ff7d2-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ff7d2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ff7d2-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ff7d2-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="ff7d2-113">成员</span><span class="sxs-lookup"><span data-stu-id="ff7d2-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ff7d2-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ff7d2-114">displayName :String</span></span>

<span data-ttu-id="ff7d2-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="ff7d2-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ff7d2-116">Type</span><span class="sxs-lookup"><span data-stu-id="ff7d2-116">Type</span></span>

*   <span data-ttu-id="ff7d2-117">String</span><span class="sxs-lookup"><span data-stu-id="ff7d2-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff7d2-118">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-118">Requirements</span></span>

|<span data-ttu-id="ff7d2-119">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-119">Requirement</span></span>| <span data-ttu-id="ff7d2-120">值</span><span class="sxs-lookup"><span data-stu-id="ff7d2-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff7d2-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ff7d2-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff7d2-122">1.0</span><span class="sxs-lookup"><span data-stu-id="ff7d2-122">1.0</span></span>|
|[<span data-ttu-id="ff7d2-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ff7d2-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff7d2-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff7d2-124">ReadItem</span></span>|
|[<span data-ttu-id="ff7d2-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ff7d2-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ff7d2-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ff7d2-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff7d2-127">示例</span><span class="sxs-lookup"><span data-stu-id="ff7d2-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ff7d2-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ff7d2-128">emailAddress :String</span></span>

<span data-ttu-id="ff7d2-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="ff7d2-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ff7d2-130">Type</span><span class="sxs-lookup"><span data-stu-id="ff7d2-130">Type</span></span>

*   <span data-ttu-id="ff7d2-131">String</span><span class="sxs-lookup"><span data-stu-id="ff7d2-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff7d2-132">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-132">Requirements</span></span>

|<span data-ttu-id="ff7d2-133">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-133">Requirement</span></span>| <span data-ttu-id="ff7d2-134">值</span><span class="sxs-lookup"><span data-stu-id="ff7d2-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff7d2-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ff7d2-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff7d2-136">1.0</span><span class="sxs-lookup"><span data-stu-id="ff7d2-136">1.0</span></span>|
|[<span data-ttu-id="ff7d2-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ff7d2-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff7d2-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff7d2-138">ReadItem</span></span>|
|[<span data-ttu-id="ff7d2-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ff7d2-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ff7d2-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ff7d2-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff7d2-141">示例</span><span class="sxs-lookup"><span data-stu-id="ff7d2-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ff7d2-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ff7d2-142">timeZone :String</span></span>

<span data-ttu-id="ff7d2-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="ff7d2-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ff7d2-144">Type</span><span class="sxs-lookup"><span data-stu-id="ff7d2-144">Type</span></span>

*   <span data-ttu-id="ff7d2-145">String</span><span class="sxs-lookup"><span data-stu-id="ff7d2-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff7d2-146">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-146">Requirements</span></span>

|<span data-ttu-id="ff7d2-147">要求</span><span class="sxs-lookup"><span data-stu-id="ff7d2-147">Requirement</span></span>| <span data-ttu-id="ff7d2-148">值</span><span class="sxs-lookup"><span data-stu-id="ff7d2-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff7d2-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ff7d2-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff7d2-150">1.0</span><span class="sxs-lookup"><span data-stu-id="ff7d2-150">1.0</span></span>|
|[<span data-ttu-id="ff7d2-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ff7d2-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff7d2-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff7d2-152">ReadItem</span></span>|
|[<span data-ttu-id="ff7d2-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ff7d2-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ff7d2-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ff7d2-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff7d2-155">示例</span><span class="sxs-lookup"><span data-stu-id="ff7d2-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
