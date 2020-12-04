---
title: Office。上下文要求集1。9
description: 使用邮箱 API 要求集1.9 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 3a8a9fe65ebf3c5a5ee63766f71dfce8e3f8d905
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570721"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="b07fd-103"> (邮箱要求集1.9 的上下文) </span><span class="sxs-lookup"><span data-stu-id="b07fd-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b07fd-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b07fd-104">[Office](office.md).context</span></span>

<span data-ttu-id="b07fd-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="b07fd-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b07fd-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="b07fd-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b07fd-107">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-107">Requirements</span></span>

|<span data-ttu-id="b07fd-108">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-108">Requirement</span></span>| <span data-ttu-id="b07fd-109">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-111">1.1</span></span>|
|[<span data-ttu-id="b07fd-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b07fd-114">属性</span><span class="sxs-lookup"><span data-stu-id="b07fd-114">Properties</span></span>

| <span data-ttu-id="b07fd-115">属性</span><span class="sxs-lookup"><span data-stu-id="b07fd-115">Property</span></span> | <span data-ttu-id="b07fd-116">型号</span><span class="sxs-lookup"><span data-stu-id="b07fd-116">Modes</span></span> | <span data-ttu-id="b07fd-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-117">Return type</span></span> | <span data-ttu-id="b07fd-118">最小值</span><span class="sxs-lookup"><span data-stu-id="b07fd-118">Minimum</span></span><br><span data-ttu-id="b07fd-119">要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b07fd-120">认证</span><span class="sxs-lookup"><span data-stu-id="b07fd-120">auth</span></span>](#auth-auth) | <span data-ttu-id="b07fd-121">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-121">Compose</span></span><br><span data-ttu-id="b07fd-122">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-122">Read</span></span> | [<span data-ttu-id="b07fd-123">Auth</span><span class="sxs-lookup"><span data-stu-id="b07fd-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-124">IdentityAPI 1。3</span><span class="sxs-lookup"><span data-stu-id="b07fd-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="b07fd-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b07fd-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b07fd-126">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-126">Compose</span></span><br><span data-ttu-id="b07fd-127">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-127">Read</span></span> | <span data-ttu-id="b07fd-128">String</span><span class="sxs-lookup"><span data-stu-id="b07fd-128">String</span></span> | [<span data-ttu-id="b07fd-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b07fd-130">过程</span><span class="sxs-lookup"><span data-stu-id="b07fd-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b07fd-131">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-131">Compose</span></span><br><span data-ttu-id="b07fd-132">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-132">Read</span></span> | [<span data-ttu-id="b07fd-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b07fd-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b07fd-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b07fd-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b07fd-136">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-136">Compose</span></span><br><span data-ttu-id="b07fd-137">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-137">Read</span></span> | <span data-ttu-id="b07fd-138">String</span><span class="sxs-lookup"><span data-stu-id="b07fd-138">String</span></span> | [<span data-ttu-id="b07fd-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b07fd-140">host</span><span class="sxs-lookup"><span data-stu-id="b07fd-140">host</span></span>](#host-hosttype) | <span data-ttu-id="b07fd-141">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-141">Compose</span></span><br><span data-ttu-id="b07fd-142">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-142">Read</span></span> | [<span data-ttu-id="b07fd-143">HostType</span><span class="sxs-lookup"><span data-stu-id="b07fd-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-144">1.5</span><span class="sxs-lookup"><span data-stu-id="b07fd-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b07fd-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="b07fd-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b07fd-146">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-146">Compose</span></span><br><span data-ttu-id="b07fd-147">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-147">Read</span></span> | [<span data-ttu-id="b07fd-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="b07fd-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b07fd-150">平台</span><span class="sxs-lookup"><span data-stu-id="b07fd-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b07fd-151">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-151">Compose</span></span><br><span data-ttu-id="b07fd-152">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-152">Read</span></span> | [<span data-ttu-id="b07fd-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b07fd-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-154">1.5</span><span class="sxs-lookup"><span data-stu-id="b07fd-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b07fd-155">满足</span><span class="sxs-lookup"><span data-stu-id="b07fd-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b07fd-156">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-156">Compose</span></span><br><span data-ttu-id="b07fd-157">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-157">Read</span></span> | [<span data-ttu-id="b07fd-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b07fd-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b07fd-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b07fd-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b07fd-161">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-161">Compose</span></span><br><span data-ttu-id="b07fd-162">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-162">Read</span></span> | [<span data-ttu-id="b07fd-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b07fd-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b07fd-165">ui</span><span class="sxs-lookup"><span data-stu-id="b07fd-165">ui</span></span>](#ui-ui) | <span data-ttu-id="b07fd-166">撰写</span><span class="sxs-lookup"><span data-stu-id="b07fd-166">Compose</span></span><br><span data-ttu-id="b07fd-167">阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-167">Read</span></span> | [<span data-ttu-id="b07fd-168">UI</span><span class="sxs-lookup"><span data-stu-id="b07fd-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b07fd-169">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b07fd-170">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="b07fd-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="b07fd-171">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="b07fd-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="b07fd-172">通过提供允许 Office 应用程序获取对加载项 web 应用程序的访问令牌的方法，支持 [单一登录 (SSO) ](../../../outlook/authenticate-a-user-with-an-sso-token.md) 。</span><span class="sxs-lookup"><span data-stu-id="b07fd-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="b07fd-173">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="b07fd-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="b07fd-174">请参阅 [IdentityAPI 1.3 要求集](../../requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="b07fd-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-175">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-175">Type</span></span>

*   [<span data-ttu-id="b07fd-176">Auth</span><span class="sxs-lookup"><span data-stu-id="b07fd-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="b07fd-177">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-177">Requirements</span></span>

|<span data-ttu-id="b07fd-178">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-178">Requirement</span></span>| <span data-ttu-id="b07fd-179">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-180">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-181">无</span><span class="sxs-lookup"><span data-stu-id="b07fd-181">N/A</span></span>|
|[<span data-ttu-id="b07fd-182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-183">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-184">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-184">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="b07fd-185">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b07fd-185">contentLanguage: String</span></span>

<span data-ttu-id="b07fd-186">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="b07fd-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b07fd-187">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言** 指定的当前 **编辑语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="b07fd-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-188">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-188">Type</span></span>

*   <span data-ttu-id="b07fd-189">String</span><span class="sxs-lookup"><span data-stu-id="b07fd-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b07fd-190">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-190">Requirements</span></span>

|<span data-ttu-id="b07fd-191">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-191">Requirement</span></span>| <span data-ttu-id="b07fd-192">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-193">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-194">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-194">1.1</span></span>|
|[<span data-ttu-id="b07fd-195">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-196">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-197">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-197">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b07fd-198">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b07fd-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b07fd-199">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="b07fd-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-200">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-200">Type</span></span>

*   [<span data-ttu-id="b07fd-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b07fd-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b07fd-202">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-202">Requirements</span></span>

|<span data-ttu-id="b07fd-203">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-203">Requirement</span></span>| <span data-ttu-id="b07fd-204">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-206">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-206">1.1</span></span>|
|[<span data-ttu-id="b07fd-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-209">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b07fd-210">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b07fd-210">displayLanguage: String</span></span>

<span data-ttu-id="b07fd-211">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="b07fd-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="b07fd-212">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的 **File > Options > 语言** 指定的当前 **显示语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="b07fd-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-213">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-213">Type</span></span>

*   <span data-ttu-id="b07fd-214">String</span><span class="sxs-lookup"><span data-stu-id="b07fd-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b07fd-215">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-215">Requirements</span></span>

|<span data-ttu-id="b07fd-216">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-216">Requirement</span></span>| <span data-ttu-id="b07fd-217">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-219">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-219">1.1</span></span>|
|[<span data-ttu-id="b07fd-220">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-221">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-222">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-222">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="b07fd-223">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b07fd-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b07fd-224">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="b07fd-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b07fd-225">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="b07fd-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-226">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-226">Type</span></span>

*   [<span data-ttu-id="b07fd-227">HostType</span><span class="sxs-lookup"><span data-stu-id="b07fd-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b07fd-228">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-228">Requirements</span></span>

|<span data-ttu-id="b07fd-229">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-229">Requirement</span></span>| <span data-ttu-id="b07fd-230">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-231">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-232">1.5</span><span class="sxs-lookup"><span data-stu-id="b07fd-232">1.5</span></span>|
|[<span data-ttu-id="b07fd-233">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-234">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-235">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b07fd-236">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b07fd-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b07fd-237">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="b07fd-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="b07fd-238">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="b07fd-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-239">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-239">Type</span></span>

*   [<span data-ttu-id="b07fd-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b07fd-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b07fd-241">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-241">Requirements</span></span>

|<span data-ttu-id="b07fd-242">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-242">Requirement</span></span>| <span data-ttu-id="b07fd-243">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-244">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-245">1.5</span><span class="sxs-lookup"><span data-stu-id="b07fd-245">1.5</span></span>|
|[<span data-ttu-id="b07fd-246">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-247">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-248">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b07fd-249">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b07fd-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b07fd-250">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="b07fd-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-251">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-251">Type</span></span>

*   [<span data-ttu-id="b07fd-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b07fd-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b07fd-253">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-253">Requirements</span></span>

|<span data-ttu-id="b07fd-254">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-254">Requirement</span></span>| <span data-ttu-id="b07fd-255">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-256">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-257">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-257">1.1</span></span>|
|[<span data-ttu-id="b07fd-258">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-259">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b07fd-260">示例</span><span class="sxs-lookup"><span data-stu-id="b07fd-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b07fd-261">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b07fd-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b07fd-262">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="b07fd-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b07fd-263">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="b07fd-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-264">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-264">Type</span></span>

*   [<span data-ttu-id="b07fd-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b07fd-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b07fd-266">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-266">Requirements</span></span>

|<span data-ttu-id="b07fd-267">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-267">Requirement</span></span>| <span data-ttu-id="b07fd-268">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-269">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-270">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-270">1.1</span></span>|
|[<span data-ttu-id="b07fd-271">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b07fd-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b07fd-272">受限</span><span class="sxs-lookup"><span data-stu-id="b07fd-272">Restricted</span></span>|
|[<span data-ttu-id="b07fd-273">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-274">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b07fd-275">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b07fd-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b07fd-276">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="b07fd-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b07fd-277">类型</span><span class="sxs-lookup"><span data-stu-id="b07fd-277">Type</span></span>

*   [<span data-ttu-id="b07fd-278">UI</span><span class="sxs-lookup"><span data-stu-id="b07fd-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b07fd-279">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-279">Requirements</span></span>

|<span data-ttu-id="b07fd-280">要求</span><span class="sxs-lookup"><span data-stu-id="b07fd-280">Requirement</span></span>| <span data-ttu-id="b07fd-281">值</span><span class="sxs-lookup"><span data-stu-id="b07fd-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="b07fd-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b07fd-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b07fd-283">1.1</span><span class="sxs-lookup"><span data-stu-id="b07fd-283">1.1</span></span>|
|[<span data-ttu-id="b07fd-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b07fd-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b07fd-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b07fd-285">Compose or Read</span></span>|
