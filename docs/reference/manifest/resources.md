---
title: 清单文件中的 Resources 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0707df137d075a9922836e5d960216d089c56675
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433899"
---
# <a name="resources-element"></a><span data-ttu-id="1954b-102">Resources 元素</span><span class="sxs-lookup"><span data-stu-id="1954b-102">Resources element</span></span>

<span data-ttu-id="1954b-p101">包含图标、字符串以及 [VersionOverrides](versionoverrides.md) 节点的 URL。清单元素通过使用资源的 **id** 来指定资源。这有助于将清单的大小保持在可管理的范围，尤其是当资源具有不同区域设置的版本时。**id** 在清单内必须是唯一的且最多可包含 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="1954b-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="1954b-107">每个资源可以具有一个或多个 **Override** 子元素以定义特定区域设置的不同资源。</span><span class="sxs-lookup"><span data-stu-id="1954b-107">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1954b-108">子元素</span><span class="sxs-lookup"><span data-stu-id="1954b-108">Child elements</span></span>

|  <span data-ttu-id="1954b-109">元素</span><span class="sxs-lookup"><span data-stu-id="1954b-109">Element</span></span> |  <span data-ttu-id="1954b-110">类型</span><span class="sxs-lookup"><span data-stu-id="1954b-110">Type</span></span>  |  <span data-ttu-id="1954b-111">说明</span><span class="sxs-lookup"><span data-stu-id="1954b-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1954b-112">Images</span><span class="sxs-lookup"><span data-stu-id="1954b-112">Images</span></span>](#images)            |  <span data-ttu-id="1954b-113">image</span><span class="sxs-lookup"><span data-stu-id="1954b-113">image</span></span>   |  <span data-ttu-id="1954b-114">提供指向图标图像的 HTTPS URL。</span><span class="sxs-lookup"><span data-stu-id="1954b-114">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="1954b-115">**Urls**</span><span class="sxs-lookup"><span data-stu-id="1954b-115">**Urls**</span></span>                |  <span data-ttu-id="1954b-116">url</span><span class="sxs-lookup"><span data-stu-id="1954b-116">url</span></span>     |  <span data-ttu-id="1954b-p102">提供 HTTPS URL 位置。一个 URL 最多可包含 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="1954b-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="1954b-119">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="1954b-119">**ShortStrings**</span></span> |  <span data-ttu-id="1954b-120">string</span><span class="sxs-lookup"><span data-stu-id="1954b-120">string</span></span>  |  <span data-ttu-id="1954b-p103">**Label** 和 **Title** 元素的文本。每个 **String** 最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="1954b-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="1954b-123">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="1954b-123">**LongStrings**</span></span>  |  <span data-ttu-id="1954b-124">字符串</span><span class="sxs-lookup"><span data-stu-id="1954b-124">string</span></span>  | <span data-ttu-id="1954b-p104">**Description** 属性的文本。每个 **String** 最多可包含 250 个字符。</span><span class="sxs-lookup"><span data-stu-id="1954b-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="1954b-127">必须对 **Image** 和 **Url** 元素中的所有 URL 使用安全套接字层 (SSL)。</span><span class="sxs-lookup"><span data-stu-id="1954b-127">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="1954b-128">图像</span><span class="sxs-lookup"><span data-stu-id="1954b-128">Images</span></span>
<span data-ttu-id="1954b-129">每个图标必须具有三个 **Images** 元素，三个强制大小的各一个元素：</span><span class="sxs-lookup"><span data-stu-id="1954b-129">Each icon must have three  **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="1954b-130">16x16</span><span class="sxs-lookup"><span data-stu-id="1954b-130">16x16</span></span>
- <span data-ttu-id="1954b-131">32x32</span><span class="sxs-lookup"><span data-stu-id="1954b-131">32x32</span></span>
- <span data-ttu-id="1954b-132">80x80</span><span class="sxs-lookup"><span data-stu-id="1954b-132">80x80</span></span>

<span data-ttu-id="1954b-133">此外还支持以下其他大小，但并不是必需的：</span><span class="sxs-lookup"><span data-stu-id="1954b-133">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="1954b-134">20x20</span><span class="sxs-lookup"><span data-stu-id="1954b-134">20x20</span></span>
- <span data-ttu-id="1954b-135">24x24</span><span class="sxs-lookup"><span data-stu-id="1954b-135">24x24</span></span>
- <span data-ttu-id="1954b-136">40x40</span><span class="sxs-lookup"><span data-stu-id="1954b-136">40x40</span></span>
- <span data-ttu-id="1954b-137">48x48</span><span class="sxs-lookup"><span data-stu-id="1954b-137">48x48</span></span>
- <span data-ttu-id="1954b-138">64x64</span><span class="sxs-lookup"><span data-stu-id="1954b-138">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="1954b-139">Outlook 需要缓存图像资源的能力，以提高性能。</span><span class="sxs-lookup"><span data-stu-id="1954b-139">Important: Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="1954b-140">为此，托管图像资源的服务器不能向响应头添加任何 CACHE-CONTROL 指令。</span><span class="sxs-lookup"><span data-stu-id="1954b-140">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="1954b-141">这将导致 Outlook 自动替代泛型或默认图像。</span><span class="sxs-lookup"><span data-stu-id="1954b-141">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="1954b-142">资源示例</span><span class="sxs-lookup"><span data-stu-id="1954b-142">Resources examples</span></span> 

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
