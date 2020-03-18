---
title: Office. context-预览要求集
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9c2c661ce870e2007bd891aee040c6b3564f7b9e
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688711"
---
# <a name="context"></a><span data-ttu-id="c952a-102">context</span><span class="sxs-lookup"><span data-stu-id="c952a-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="c952a-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c952a-103">[Office](office.md).context</span></span>

<span data-ttu-id="c952a-104">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="c952a-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c952a-105">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-preview)"。</span><span class="sxs-lookup"><span data-stu-id="c952a-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c952a-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-106">Requirements</span></span>

|<span data-ttu-id="c952a-107">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-107">Requirement</span></span>| <span data-ttu-id="c952a-108">值</span><span class="sxs-lookup"><span data-stu-id="c952a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-110">1.1</span></span>|
|[<span data-ttu-id="c952a-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c952a-113">属性</span><span class="sxs-lookup"><span data-stu-id="c952a-113">Properties</span></span>

| <span data-ttu-id="c952a-114">属性</span><span class="sxs-lookup"><span data-stu-id="c952a-114">Property</span></span> | <span data-ttu-id="c952a-115">型号</span><span class="sxs-lookup"><span data-stu-id="c952a-115">Modes</span></span> | <span data-ttu-id="c952a-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="c952a-116">Return type</span></span> | <span data-ttu-id="c952a-117">最低</span><span class="sxs-lookup"><span data-stu-id="c952a-117">Minimum</span></span><br><span data-ttu-id="c952a-118">要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c952a-119">认证</span><span class="sxs-lookup"><span data-stu-id="c952a-119">auth</span></span>](#auth-auth) | <span data-ttu-id="c952a-120">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-120">Compose</span></span><br><span data-ttu-id="c952a-121">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-121">Read</span></span> | [<span data-ttu-id="c952a-122">Auth</span><span class="sxs-lookup"><span data-stu-id="c952a-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="c952a-123">预览</span><span class="sxs-lookup"><span data-stu-id="c952a-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="c952a-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c952a-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c952a-125">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-125">Compose</span></span><br><span data-ttu-id="c952a-126">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-126">Read</span></span> | <span data-ttu-id="c952a-127">String</span><span class="sxs-lookup"><span data-stu-id="c952a-127">String</span></span> | [<span data-ttu-id="c952a-128">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-129">过程</span><span class="sxs-lookup"><span data-stu-id="c952a-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c952a-130">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-130">Compose</span></span><br><span data-ttu-id="c952a-131">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-131">Read</span></span> | [<span data-ttu-id="c952a-132">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c952a-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="c952a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c952a-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c952a-135">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-135">Compose</span></span><br><span data-ttu-id="c952a-136">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-136">Read</span></span> | <span data-ttu-id="c952a-137">String</span><span class="sxs-lookup"><span data-stu-id="c952a-137">String</span></span> | [<span data-ttu-id="c952a-138">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-139">host</span><span class="sxs-lookup"><span data-stu-id="c952a-139">host</span></span>](#host-hosttype) | <span data-ttu-id="c952a-140">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-140">Compose</span></span><br><span data-ttu-id="c952a-141">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-141">Read</span></span> | [<span data-ttu-id="c952a-142">HostType</span><span class="sxs-lookup"><span data-stu-id="c952a-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="c952a-143">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="c952a-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c952a-145">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-145">Compose</span></span><br><span data-ttu-id="c952a-146">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-146">Read</span></span> | [<span data-ttu-id="c952a-147">邮箱</span><span class="sxs-lookup"><span data-stu-id="c952a-147">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="c952a-148">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="c952a-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="c952a-150">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-150">Compose</span></span><br><span data-ttu-id="c952a-151">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-151">Read</span></span> | [<span data-ttu-id="c952a-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="c952a-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="c952a-153">预览</span><span class="sxs-lookup"><span data-stu-id="c952a-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="c952a-154">平台</span><span class="sxs-lookup"><span data-stu-id="c952a-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="c952a-155">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-155">Compose</span></span><br><span data-ttu-id="c952a-156">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-156">Read</span></span> | [<span data-ttu-id="c952a-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c952a-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="c952a-158">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-159">满足</span><span class="sxs-lookup"><span data-stu-id="c952a-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c952a-160">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-160">Compose</span></span><br><span data-ttu-id="c952a-161">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-161">Read</span></span> | [<span data-ttu-id="c952a-162">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c952a-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="c952a-163">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c952a-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c952a-165">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-165">Compose</span></span><br><span data-ttu-id="c952a-166">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-166">Read</span></span> | [<span data-ttu-id="c952a-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c952a-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="c952a-168">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c952a-169">ui</span><span class="sxs-lookup"><span data-stu-id="c952a-169">ui</span></span>](#ui-ui) | <span data-ttu-id="c952a-170">撰写</span><span class="sxs-lookup"><span data-stu-id="c952a-170">Compose</span></span><br><span data-ttu-id="c952a-171">读取</span><span class="sxs-lookup"><span data-stu-id="c952a-171">Read</span></span> | [<span data-ttu-id="c952a-172">UI</span><span class="sxs-lookup"><span data-stu-id="c952a-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="c952a-173">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c952a-174">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="c952a-174">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="c952a-175">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="c952a-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="c952a-176">通过提供一种方法来支持[单一登录（SSO）](../../../outlook/authenticate-a-user-with-an-sso-token.md) ，使 Office 主机能够获取加载项的 web 应用程序的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="c952a-176">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="c952a-177">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="c952a-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-178">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-178">Type</span></span>

*   [<span data-ttu-id="c952a-179">Auth</span><span class="sxs-lookup"><span data-stu-id="c952a-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="c952a-180">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-180">Requirements</span></span>

|<span data-ttu-id="c952a-181">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-181">Requirement</span></span>| <span data-ttu-id="c952a-182">值</span><span class="sxs-lookup"><span data-stu-id="c952a-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-183">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-184">预览</span><span class="sxs-lookup"><span data-stu-id="c952a-184">Preview</span></span>|
|[<span data-ttu-id="c952a-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-185">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-187">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="c952a-188">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="c952a-188">contentLanguage: String</span></span>

<span data-ttu-id="c952a-189">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="c952a-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c952a-190">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="c952a-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-191">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-191">Type</span></span>

*   <span data-ttu-id="c952a-192">String</span><span class="sxs-lookup"><span data-stu-id="c952a-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c952a-193">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-193">Requirements</span></span>

|<span data-ttu-id="c952a-194">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-194">Requirement</span></span>| <span data-ttu-id="c952a-195">值</span><span class="sxs-lookup"><span data-stu-id="c952a-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-196">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-197">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-197">1.1</span></span>|
|[<span data-ttu-id="c952a-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-198">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-200">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-200">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c952a-201">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c952a-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c952a-202">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="c952a-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-203">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-203">Type</span></span>

*   [<span data-ttu-id="c952a-204">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c952a-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c952a-205">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-205">Requirements</span></span>

|<span data-ttu-id="c952a-206">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-206">Requirement</span></span>| <span data-ttu-id="c952a-207">值</span><span class="sxs-lookup"><span data-stu-id="c952a-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-208">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-209">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-209">1.1</span></span>|
|[<span data-ttu-id="c952a-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-210">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-212">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c952a-213">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="c952a-213">displayLanguage: String</span></span>

<span data-ttu-id="c952a-214">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="c952a-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c952a-215">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="c952a-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-216">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-216">Type</span></span>

*   <span data-ttu-id="c952a-217">String</span><span class="sxs-lookup"><span data-stu-id="c952a-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c952a-218">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-218">Requirements</span></span>

|<span data-ttu-id="c952a-219">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-219">Requirement</span></span>| <span data-ttu-id="c952a-220">值</span><span class="sxs-lookup"><span data-stu-id="c952a-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-221">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-222">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-222">1.1</span></span>|
|[<span data-ttu-id="c952a-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-224">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-225">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-225">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="c952a-226">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="c952a-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="c952a-227">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="c952a-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-228">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-228">Type</span></span>

*   [<span data-ttu-id="c952a-229">HostType</span><span class="sxs-lookup"><span data-stu-id="c952a-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="c952a-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-230">Requirements</span></span>

|<span data-ttu-id="c952a-231">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-231">Requirement</span></span>| <span data-ttu-id="c952a-232">值</span><span class="sxs-lookup"><span data-stu-id="c952a-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-234">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-234">1.1</span></span>|
|[<span data-ttu-id="c952a-235">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-236">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-237">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="c952a-238">officeTheme： [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="c952a-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="c952a-239">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="c952a-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="c952a-240">此成员仅在 Windows 中的 Outlook 中受支持。</span><span class="sxs-lookup"><span data-stu-id="c952a-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="c952a-241">使用 Office 主题颜色，您可以将加载项的配色方案与用户选择的当前 Office 主题进行协调，以供用户使用**office > Office 帐户 > Office 主题 UI**，该用户在所有 Office 主机应用程序中应用。</span><span class="sxs-lookup"><span data-stu-id="c952a-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="c952a-242">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="c952a-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-243">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-243">Type</span></span>

*   [<span data-ttu-id="c952a-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="c952a-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="c952a-245">属性：</span><span class="sxs-lookup"><span data-stu-id="c952a-245">Properties:</span></span>

|<span data-ttu-id="c952a-246">姓名</span><span class="sxs-lookup"><span data-stu-id="c952a-246">Name</span></span>| <span data-ttu-id="c952a-247">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-247">Type</span></span>| <span data-ttu-id="c952a-248">说明</span><span class="sxs-lookup"><span data-stu-id="c952a-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="c952a-249">String</span><span class="sxs-lookup"><span data-stu-id="c952a-249">String</span></span>|<span data-ttu-id="c952a-250">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="c952a-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="c952a-251">String</span><span class="sxs-lookup"><span data-stu-id="c952a-251">String</span></span>|<span data-ttu-id="c952a-252">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="c952a-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="c952a-253">String</span><span class="sxs-lookup"><span data-stu-id="c952a-253">String</span></span>|<span data-ttu-id="c952a-254">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="c952a-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="c952a-255">字符串</span><span class="sxs-lookup"><span data-stu-id="c952a-255">String</span></span>|<span data-ttu-id="c952a-256">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="c952a-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c952a-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-257">Requirements</span></span>

|<span data-ttu-id="c952a-258">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-258">Requirement</span></span>| <span data-ttu-id="c952a-259">值</span><span class="sxs-lookup"><span data-stu-id="c952a-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-261">预览</span><span class="sxs-lookup"><span data-stu-id="c952a-261">Preview</span></span>|
|[<span data-ttu-id="c952a-262">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-263">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-264">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-264">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="c952a-265">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="c952a-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="c952a-266">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="c952a-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-267">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-267">Type</span></span>

*   [<span data-ttu-id="c952a-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c952a-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="c952a-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-269">Requirements</span></span>

|<span data-ttu-id="c952a-270">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-270">Requirement</span></span>| <span data-ttu-id="c952a-271">值</span><span class="sxs-lookup"><span data-stu-id="c952a-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-272">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-273">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-273">1.1</span></span>|
|[<span data-ttu-id="c952a-274">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-274">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-275">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-276">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c952a-277">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c952a-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c952a-278">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="c952a-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-279">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-279">Type</span></span>

*   [<span data-ttu-id="c952a-280">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c952a-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c952a-281">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-281">Requirements</span></span>

|<span data-ttu-id="c952a-282">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-282">Requirement</span></span>| <span data-ttu-id="c952a-283">值</span><span class="sxs-lookup"><span data-stu-id="c952a-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-285">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-285">1.1</span></span>|
|[<span data-ttu-id="c952a-286">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-286">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-287">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c952a-288">示例</span><span class="sxs-lookup"><span data-stu-id="c952a-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c952a-289">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c952a-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c952a-290">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="c952a-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c952a-291">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="c952a-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-292">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-292">Type</span></span>

*   [<span data-ttu-id="c952a-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c952a-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c952a-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-294">Requirements</span></span>

|<span data-ttu-id="c952a-295">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-295">Requirement</span></span>| <span data-ttu-id="c952a-296">值</span><span class="sxs-lookup"><span data-stu-id="c952a-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-297">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-298">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-298">1.1</span></span>|
|[<span data-ttu-id="c952a-299">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c952a-299">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c952a-300">受限</span><span class="sxs-lookup"><span data-stu-id="c952a-300">Restricted</span></span>|
|[<span data-ttu-id="c952a-301">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-301">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-302">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c952a-303">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c952a-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c952a-304">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="c952a-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c952a-305">类型</span><span class="sxs-lookup"><span data-stu-id="c952a-305">Type</span></span>

*   [<span data-ttu-id="c952a-306">UI</span><span class="sxs-lookup"><span data-stu-id="c952a-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c952a-307">Requirements</span><span class="sxs-lookup"><span data-stu-id="c952a-307">Requirements</span></span>

|<span data-ttu-id="c952a-308">要求</span><span class="sxs-lookup"><span data-stu-id="c952a-308">Requirement</span></span>| <span data-ttu-id="c952a-309">值</span><span class="sxs-lookup"><span data-stu-id="c952a-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="c952a-310">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c952a-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c952a-311">1.1</span><span class="sxs-lookup"><span data-stu-id="c952a-311">1.1</span></span>|
|[<span data-ttu-id="c952a-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c952a-312">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c952a-313">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c952a-313">Compose or Read</span></span>|
