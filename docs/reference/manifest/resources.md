---
title: 清单文件中的 Resources 元素
description: Resources 元素包含用于 VersionOverrides 节点的图标、字符串和 URL。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0a528b05904ef65c3643aaebb9149eb2091e2287
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641268"
---
# <a name="resources-element"></a><span data-ttu-id="fd7f8-103">Resources 元素</span><span class="sxs-lookup"><span data-stu-id="fd7f8-103">Resources element</span></span>

<span data-ttu-id="fd7f8-p101">包含图标、字符串以及 [VersionOverrides](versionoverrides.md) 节点的 URL。清单元素通过使用资源的 **id** 来指定资源。这有助于将清单的大小保持在可管理的范围，尤其是当资源具有不同区域设置的版本时。**id** 在清单内必须是唯一的且最多可包含 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="fd7f8-108">每个资源可以具有一个或多个 **Override** 子元素以定义特定区域设置的不同资源。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-108">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="fd7f8-109">子元素</span><span class="sxs-lookup"><span data-stu-id="fd7f8-109">Child elements</span></span>

|  <span data-ttu-id="fd7f8-110">元素</span><span class="sxs-lookup"><span data-stu-id="fd7f8-110">Element</span></span> |  <span data-ttu-id="fd7f8-111">类型</span><span class="sxs-lookup"><span data-stu-id="fd7f8-111">Type</span></span>  |  <span data-ttu-id="fd7f8-112">说明</span><span class="sxs-lookup"><span data-stu-id="fd7f8-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fd7f8-113">Images</span><span class="sxs-lookup"><span data-stu-id="fd7f8-113">Images</span></span>](#images)            |  <span data-ttu-id="fd7f8-114">image</span><span class="sxs-lookup"><span data-stu-id="fd7f8-114">image</span></span>   |  <span data-ttu-id="fd7f8-115">提供指向图标图像的 HTTPS URL。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-115">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="fd7f8-116">**Urls**</span><span class="sxs-lookup"><span data-stu-id="fd7f8-116">**Urls**</span></span>                |  <span data-ttu-id="fd7f8-117">url</span><span class="sxs-lookup"><span data-stu-id="fd7f8-117">url</span></span>     |  <span data-ttu-id="fd7f8-p102">提供 HTTPS URL 位置。一个 URL 最多可包含 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="fd7f8-120">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="fd7f8-120">**ShortStrings**</span></span> |  <span data-ttu-id="fd7f8-121">string</span><span class="sxs-lookup"><span data-stu-id="fd7f8-121">string</span></span>  |  <span data-ttu-id="fd7f8-p103">**Label** 和 **Title** 元素的文本。每个 **String** 最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="fd7f8-124">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="fd7f8-124">**LongStrings**</span></span>  |  <span data-ttu-id="fd7f8-125">string</span><span class="sxs-lookup"><span data-stu-id="fd7f8-125">string</span></span>  | <span data-ttu-id="fd7f8-p104">**Description** 属性的文本。每个 **String** 最多可包含 250 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="fd7f8-128">必须对 **Image** 和 **Url** 元素中的所有 URL 使用安全套接字层 (SSL)。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-128">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="fd7f8-129">图像</span><span class="sxs-lookup"><span data-stu-id="fd7f8-129">Images</span></span>
<span data-ttu-id="fd7f8-130">每个图标必须具有三个**Images**元素，三个强制大小的元素分别为：</span><span class="sxs-lookup"><span data-stu-id="fd7f8-130">Each icon must have three **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="fd7f8-131">16x16</span><span class="sxs-lookup"><span data-stu-id="fd7f8-131">16x16</span></span>
- <span data-ttu-id="fd7f8-132">32x32</span><span class="sxs-lookup"><span data-stu-id="fd7f8-132">32x32</span></span>
- <span data-ttu-id="fd7f8-133">80x80</span><span class="sxs-lookup"><span data-stu-id="fd7f8-133">80x80</span></span>

<span data-ttu-id="fd7f8-134">此外还支持以下其他大小，但并不是必需的：</span><span class="sxs-lookup"><span data-stu-id="fd7f8-134">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="fd7f8-135">20x20</span><span class="sxs-lookup"><span data-stu-id="fd7f8-135">20x20</span></span>
- <span data-ttu-id="fd7f8-136">24x24</span><span class="sxs-lookup"><span data-stu-id="fd7f8-136">24x24</span></span>
- <span data-ttu-id="fd7f8-137">40x40</span><span class="sxs-lookup"><span data-stu-id="fd7f8-137">40x40</span></span>
- <span data-ttu-id="fd7f8-138">48x48</span><span class="sxs-lookup"><span data-stu-id="fd7f8-138">48x48</span></span>
- <span data-ttu-id="fd7f8-139">64x64</span><span class="sxs-lookup"><span data-stu-id="fd7f8-139">64x64</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fd7f8-140">Outlook 需要缓存图像资源的能力，以提高性能。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-140">Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="fd7f8-141">为此，托管图像资源的服务器不能向响应头添加任何 CACHE-CONTROL 指令。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-141">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="fd7f8-142">这将导致 Outlook 自动替代泛型或默认图像。</span><span class="sxs-lookup"><span data-stu-id="fd7f8-142">This will result in Outlook automatically substituting a generic or default image.</span></span>

## <a name="resources-examples"></a><span data-ttu-id="fd7f8-143">资源示例</span><span class="sxs-lookup"><span data-stu-id="fd7f8-143">Resources examples</span></span>

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
