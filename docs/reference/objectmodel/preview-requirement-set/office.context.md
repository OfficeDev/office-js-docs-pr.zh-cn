---
title: Office. context-预览要求集
description: 使用邮箱 API preview 要求集的 Outlook 外接程序可用的 Context 对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8370df907aa3ab0534254057860c187cec583e6c
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570784"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="12f53-103"> (邮箱预览要求集的上下文) </span><span class="sxs-lookup"><span data-stu-id="12f53-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="12f53-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="12f53-104">[Office](office.md).context</span></span>

<span data-ttu-id="12f53-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="12f53-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="12f53-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="12f53-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f53-107">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-107">Requirements</span></span>

|<span data-ttu-id="12f53-108">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-108">Requirement</span></span>| <span data-ttu-id="12f53-109">值</span><span class="sxs-lookup"><span data-stu-id="12f53-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-111">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-111">1.1</span></span>|
|[<span data-ttu-id="12f53-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="12f53-114">属性</span><span class="sxs-lookup"><span data-stu-id="12f53-114">Properties</span></span>

| <span data-ttu-id="12f53-115">属性</span><span class="sxs-lookup"><span data-stu-id="12f53-115">Property</span></span> | <span data-ttu-id="12f53-116">型号</span><span class="sxs-lookup"><span data-stu-id="12f53-116">Modes</span></span> | <span data-ttu-id="12f53-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="12f53-117">Return type</span></span> | <span data-ttu-id="12f53-118">最小值</span><span class="sxs-lookup"><span data-stu-id="12f53-118">Minimum</span></span><br><span data-ttu-id="12f53-119">要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="12f53-120">认证</span><span class="sxs-lookup"><span data-stu-id="12f53-120">auth</span></span>](#auth-auth) | <span data-ttu-id="12f53-121">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-121">Compose</span></span><br><span data-ttu-id="12f53-122">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-122">Read</span></span> | [<span data-ttu-id="12f53-123">Auth</span><span class="sxs-lookup"><span data-stu-id="12f53-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-124">IdentityAPI 1。3</span><span class="sxs-lookup"><span data-stu-id="12f53-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="12f53-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="12f53-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="12f53-126">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-126">Compose</span></span><br><span data-ttu-id="12f53-127">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-127">Read</span></span> | <span data-ttu-id="12f53-128">String</span><span class="sxs-lookup"><span data-stu-id="12f53-128">String</span></span> | [<span data-ttu-id="12f53-129">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="12f53-130">过程</span><span class="sxs-lookup"><span data-stu-id="12f53-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="12f53-131">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-131">Compose</span></span><br><span data-ttu-id="12f53-132">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-132">Read</span></span> | [<span data-ttu-id="12f53-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="12f53-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-134">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="12f53-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="12f53-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="12f53-136">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-136">Compose</span></span><br><span data-ttu-id="12f53-137">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-137">Read</span></span> | <span data-ttu-id="12f53-138">String</span><span class="sxs-lookup"><span data-stu-id="12f53-138">String</span></span> | [<span data-ttu-id="12f53-139">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="12f53-140">host</span><span class="sxs-lookup"><span data-stu-id="12f53-140">host</span></span>](#host-hosttype) | <span data-ttu-id="12f53-141">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-141">Compose</span></span><br><span data-ttu-id="12f53-142">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-142">Read</span></span> | [<span data-ttu-id="12f53-143">HostType</span><span class="sxs-lookup"><span data-stu-id="12f53-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-144">1.5</span><span class="sxs-lookup"><span data-stu-id="12f53-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="12f53-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="12f53-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="12f53-146">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-146">Compose</span></span><br><span data-ttu-id="12f53-147">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-147">Read</span></span> | [<span data-ttu-id="12f53-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="12f53-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-149">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="12f53-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="12f53-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="12f53-151">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-151">Compose</span></span><br><span data-ttu-id="12f53-152">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-152">Read</span></span> | [<span data-ttu-id="12f53-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="12f53-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-154">预览</span><span class="sxs-lookup"><span data-stu-id="12f53-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="12f53-155">平台</span><span class="sxs-lookup"><span data-stu-id="12f53-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="12f53-156">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-156">Compose</span></span><br><span data-ttu-id="12f53-157">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-157">Read</span></span> | [<span data-ttu-id="12f53-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="12f53-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-159">1.5</span><span class="sxs-lookup"><span data-stu-id="12f53-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="12f53-160">满足</span><span class="sxs-lookup"><span data-stu-id="12f53-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="12f53-161">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-161">Compose</span></span><br><span data-ttu-id="12f53-162">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-162">Read</span></span> | [<span data-ttu-id="12f53-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="12f53-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-164">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="12f53-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="12f53-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="12f53-166">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-166">Compose</span></span><br><span data-ttu-id="12f53-167">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-167">Read</span></span> | [<span data-ttu-id="12f53-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="12f53-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-169">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="12f53-170">ui</span><span class="sxs-lookup"><span data-stu-id="12f53-170">ui</span></span>](#ui-ui) | <span data-ttu-id="12f53-171">撰写</span><span class="sxs-lookup"><span data-stu-id="12f53-171">Compose</span></span><br><span data-ttu-id="12f53-172">阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-172">Read</span></span> | [<span data-ttu-id="12f53-173">UI</span><span class="sxs-lookup"><span data-stu-id="12f53-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="12f53-174">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="12f53-175">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="12f53-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="12f53-176">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="12f53-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="12f53-177">通过提供允许 Office 应用程序获取对加载项 web 应用程序的访问令牌的方法，支持 [单一登录 (SSO) ](../../../outlook/authenticate-a-user-with-an-sso-token.md) 。</span><span class="sxs-lookup"><span data-stu-id="12f53-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="12f53-178">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="12f53-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-179">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-179">Type</span></span>

*   [<span data-ttu-id="12f53-180">Auth</span><span class="sxs-lookup"><span data-stu-id="12f53-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="12f53-181">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-181">Requirements</span></span>

|<span data-ttu-id="12f53-182">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-182">Requirement</span></span>| <span data-ttu-id="12f53-183">值</span><span class="sxs-lookup"><span data-stu-id="12f53-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-185">预览</span><span class="sxs-lookup"><span data-stu-id="12f53-185">Preview</span></span>|
|[<span data-ttu-id="12f53-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-188">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="12f53-189">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="12f53-189">contentLanguage: String</span></span>

<span data-ttu-id="12f53-190">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="12f53-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="12f53-191">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言** 指定的当前 **编辑语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="12f53-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-192">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-192">Type</span></span>

*   <span data-ttu-id="12f53-193">String</span><span class="sxs-lookup"><span data-stu-id="12f53-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f53-194">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-194">Requirements</span></span>

|<span data-ttu-id="12f53-195">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-195">Requirement</span></span>| <span data-ttu-id="12f53-196">值</span><span class="sxs-lookup"><span data-stu-id="12f53-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-198">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-198">1.1</span></span>|
|[<span data-ttu-id="12f53-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-201">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="12f53-202">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="12f53-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="12f53-203">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="12f53-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-204">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-204">Type</span></span>

*   [<span data-ttu-id="12f53-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="12f53-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="12f53-206">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-206">Requirements</span></span>

|<span data-ttu-id="12f53-207">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-207">Requirement</span></span>| <span data-ttu-id="12f53-208">值</span><span class="sxs-lookup"><span data-stu-id="12f53-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-210">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-210">1.1</span></span>|
|[<span data-ttu-id="12f53-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-212">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-213">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="12f53-214">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="12f53-214">displayLanguage: String</span></span>

<span data-ttu-id="12f53-215">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="12f53-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="12f53-216">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的 **File > Options > 语言** 指定的当前 **显示语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="12f53-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-217">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-217">Type</span></span>

*   <span data-ttu-id="12f53-218">String</span><span class="sxs-lookup"><span data-stu-id="12f53-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f53-219">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-219">Requirements</span></span>

|<span data-ttu-id="12f53-220">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-220">Requirement</span></span>| <span data-ttu-id="12f53-221">值</span><span class="sxs-lookup"><span data-stu-id="12f53-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-223">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-223">1.1</span></span>|
|[<span data-ttu-id="12f53-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-226">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="12f53-227">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="12f53-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="12f53-228">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="12f53-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="12f53-229">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取主机。</span><span class="sxs-lookup"><span data-stu-id="12f53-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-230">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-230">Type</span></span>

*   [<span data-ttu-id="12f53-231">HostType</span><span class="sxs-lookup"><span data-stu-id="12f53-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="12f53-232">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-232">Requirements</span></span>

|<span data-ttu-id="12f53-233">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-233">Requirement</span></span>| <span data-ttu-id="12f53-234">值</span><span class="sxs-lookup"><span data-stu-id="12f53-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-236">1.5</span><span class="sxs-lookup"><span data-stu-id="12f53-236">1.5</span></span>|
|[<span data-ttu-id="12f53-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-239">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="12f53-240">officeTheme： [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="12f53-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="12f53-241">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="12f53-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="12f53-242">此成员仅在 Windows 中的 Outlook 中受支持。</span><span class="sxs-lookup"><span data-stu-id="12f53-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="12f53-243">使用 Office 主题颜色，您可以将加载项的配色方案与用户选择的当前 Office 主题进行协调，以供用户使用 **office > Office 帐户 > Office 主题 UI**，该用户在所有 Office 客户端应用程序中应用。</span><span class="sxs-lookup"><span data-stu-id="12f53-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="12f53-244">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="12f53-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-245">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-245">Type</span></span>

*   [<span data-ttu-id="12f53-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="12f53-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="12f53-247">属性：</span><span class="sxs-lookup"><span data-stu-id="12f53-247">Properties:</span></span>

|<span data-ttu-id="12f53-248">名称</span><span class="sxs-lookup"><span data-stu-id="12f53-248">Name</span></span>| <span data-ttu-id="12f53-249">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-249">Type</span></span>| <span data-ttu-id="12f53-250">说明</span><span class="sxs-lookup"><span data-stu-id="12f53-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="12f53-251">String</span><span class="sxs-lookup"><span data-stu-id="12f53-251">String</span></span>|<span data-ttu-id="12f53-252">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="12f53-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="12f53-253">String</span><span class="sxs-lookup"><span data-stu-id="12f53-253">String</span></span>|<span data-ttu-id="12f53-254">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="12f53-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="12f53-255">字符串</span><span class="sxs-lookup"><span data-stu-id="12f53-255">String</span></span>|<span data-ttu-id="12f53-256">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="12f53-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="12f53-257">字符串</span><span class="sxs-lookup"><span data-stu-id="12f53-257">String</span></span>|<span data-ttu-id="12f53-258">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="12f53-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12f53-259">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-259">Requirements</span></span>

|<span data-ttu-id="12f53-260">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-260">Requirement</span></span>| <span data-ttu-id="12f53-261">值</span><span class="sxs-lookup"><span data-stu-id="12f53-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-262">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-263">预览</span><span class="sxs-lookup"><span data-stu-id="12f53-263">Preview</span></span>|
|[<span data-ttu-id="12f53-264">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-265">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-266">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-266">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="12f53-267">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="12f53-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="12f53-268">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="12f53-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="12f53-269">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="12f53-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-270">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-270">Type</span></span>

*   [<span data-ttu-id="12f53-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="12f53-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="12f53-272">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-272">Requirements</span></span>

|<span data-ttu-id="12f53-273">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-273">Requirement</span></span>| <span data-ttu-id="12f53-274">值</span><span class="sxs-lookup"><span data-stu-id="12f53-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-275">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-276">1.5</span><span class="sxs-lookup"><span data-stu-id="12f53-276">1.5</span></span>|
|[<span data-ttu-id="12f53-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-278">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-279">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="12f53-280">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="12f53-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="12f53-281">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="12f53-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-282">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-282">Type</span></span>

*   [<span data-ttu-id="12f53-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="12f53-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="12f53-284">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-284">Requirements</span></span>

|<span data-ttu-id="12f53-285">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-285">Requirement</span></span>| <span data-ttu-id="12f53-286">值</span><span class="sxs-lookup"><span data-stu-id="12f53-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-287">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-288">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-288">1.1</span></span>|
|[<span data-ttu-id="12f53-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f53-291">示例</span><span class="sxs-lookup"><span data-stu-id="12f53-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="12f53-292">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="12f53-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="12f53-293">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="12f53-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="12f53-294">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="12f53-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-295">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-295">Type</span></span>

*   [<span data-ttu-id="12f53-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="12f53-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="12f53-297">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-297">Requirements</span></span>

|<span data-ttu-id="12f53-298">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-298">Requirement</span></span>| <span data-ttu-id="12f53-299">值</span><span class="sxs-lookup"><span data-stu-id="12f53-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-300">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-301">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-301">1.1</span></span>|
|[<span data-ttu-id="12f53-302">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="12f53-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="12f53-303">受限</span><span class="sxs-lookup"><span data-stu-id="12f53-303">Restricted</span></span>|
|[<span data-ttu-id="12f53-304">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-305">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="12f53-306">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="12f53-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="12f53-307">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="12f53-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="12f53-308">类型</span><span class="sxs-lookup"><span data-stu-id="12f53-308">Type</span></span>

*   [<span data-ttu-id="12f53-309">UI</span><span class="sxs-lookup"><span data-stu-id="12f53-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="12f53-310">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-310">Requirements</span></span>

|<span data-ttu-id="12f53-311">要求</span><span class="sxs-lookup"><span data-stu-id="12f53-311">Requirement</span></span>| <span data-ttu-id="12f53-312">值</span><span class="sxs-lookup"><span data-stu-id="12f53-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f53-313">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="12f53-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="12f53-314">1.1</span><span class="sxs-lookup"><span data-stu-id="12f53-314">1.1</span></span>|
|[<span data-ttu-id="12f53-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="12f53-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="12f53-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="12f53-316">Compose or Read</span></span>|
