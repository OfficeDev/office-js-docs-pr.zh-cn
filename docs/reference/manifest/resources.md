# <a name="resources-element"></a><span data-ttu-id="04666-101">Resources 元素</span><span class="sxs-lookup"><span data-stu-id="04666-101">Resources element</span></span>

<span data-ttu-id="04666-p101">包含图标、字符串以及 [VersionOverrides](versionoverrides.md) 节点的 URL。清单元素通过使用资源的 **id** 来指定资源。这有助于将清单的大小保持在可管理的范围，尤其是当资源具有不同区域设置的版本时。**id** 在清单内必须是唯一的且最多可包含 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="04666-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="04666-106">每个资源可以具有一个或多个 **Override** 子元素以定义特定区域设置的不同资源。</span><span class="sxs-lookup"><span data-stu-id="04666-106">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="04666-107">子元素</span><span class="sxs-lookup"><span data-stu-id="04666-107">Child elements</span></span>

|  <span data-ttu-id="04666-108">元素</span><span class="sxs-lookup"><span data-stu-id="04666-108">Element</span></span> |  <span data-ttu-id="04666-109">类型</span><span class="sxs-lookup"><span data-stu-id="04666-109">Type</span></span>  |  <span data-ttu-id="04666-110">说明</span><span class="sxs-lookup"><span data-stu-id="04666-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="04666-111">图像</span><span class="sxs-lookup"><span data-stu-id="04666-111">Images</span></span>](#images)            |  <span data-ttu-id="04666-112">image</span><span class="sxs-lookup"><span data-stu-id="04666-112">image</span></span>   |  <span data-ttu-id="04666-113">提供指向图标图像的 HTTPS URL。</span><span class="sxs-lookup"><span data-stu-id="04666-113">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="04666-114">**Url**</span><span class="sxs-lookup"><span data-stu-id="04666-114">**Urls**</span></span>                |  <span data-ttu-id="04666-115">url</span><span class="sxs-lookup"><span data-stu-id="04666-115">url</span></span>     |  <span data-ttu-id="04666-p102">提供 HTTPS URL 位置。一个 URL 最多可包含 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="04666-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="04666-118">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="04666-118">**ShortStrings**</span></span> |  <span data-ttu-id="04666-119">string</span><span class="sxs-lookup"><span data-stu-id="04666-119">string</span></span>  |  <span data-ttu-id="04666-p103">**Label** 和 **Title** 元素的文本。每个 **String** 最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="04666-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="04666-122">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="04666-122">**LongStrings**</span></span>  |  <span data-ttu-id="04666-123">string</span><span class="sxs-lookup"><span data-stu-id="04666-123">string</span></span>  | <span data-ttu-id="04666-p104">**Description** 属性的文本。每个 **String** 最多可包含 250 个字符。</span><span class="sxs-lookup"><span data-stu-id="04666-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="04666-126">必须对 **Image** 和 **Url** 元素中的所有 URL 使用安全套接字层 (SSL)。</span><span class="sxs-lookup"><span data-stu-id="04666-126">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="04666-127">图像</span><span class="sxs-lookup"><span data-stu-id="04666-127">Images</span></span>
<span data-ttu-id="04666-128">每个图标必须具有三个 **Images** 元素，三个强制大小的各一个元素：</span><span class="sxs-lookup"><span data-stu-id="04666-128">Each icon must have three  **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="04666-129">16x16</span><span class="sxs-lookup"><span data-stu-id="04666-129">16x16</span></span>
- <span data-ttu-id="04666-130">32x32</span><span class="sxs-lookup"><span data-stu-id="04666-130">32x32</span></span>
- <span data-ttu-id="04666-131">80x80</span><span class="sxs-lookup"><span data-stu-id="04666-131">80x80</span></span>

<span data-ttu-id="04666-132">此外还支持以下其他大小，但并不是必需的：</span><span class="sxs-lookup"><span data-stu-id="04666-132">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="04666-133">20x20</span><span class="sxs-lookup"><span data-stu-id="04666-133">20x20</span></span>
- <span data-ttu-id="04666-134">24x24</span><span class="sxs-lookup"><span data-stu-id="04666-134">24x24</span></span>
- <span data-ttu-id="04666-135">40x40</span><span class="sxs-lookup"><span data-stu-id="04666-135">40x40</span></span>
- <span data-ttu-id="04666-136">48x48</span><span class="sxs-lookup"><span data-stu-id="04666-136">48x48</span></span>
- <span data-ttu-id="04666-137">64x64</span><span class="sxs-lookup"><span data-stu-id="04666-137">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="04666-138">Outlook 需要缓存图像资源的能力，以提高性能。</span><span class="sxs-lookup"><span data-stu-id="04666-138">Important: Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="04666-139">为此，托管图像资源的服务器不能向响应头添加任何 CACHE-CONTROL 指令。</span><span class="sxs-lookup"><span data-stu-id="04666-139">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="04666-140">这将导致 Outlook 自动替代泛型或默认图像。</span><span class="sxs-lookup"><span data-stu-id="04666-140">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="04666-141">资源示例</span><span class="sxs-lookup"><span data-stu-id="04666-141">Resources examples</span></span> 

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
