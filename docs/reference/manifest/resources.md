---
title: 清单文件中的 Resources 元素
description: Resources 元素包含用于 VersionOverrides 节点的图标、字符串和 URL。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 717e3cecd32fbf2bdb806f7484cc954a86b82e3d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608745"
---
# <a name="resources-element"></a>Resources 元素

包含图标、字符串以及 [VersionOverrides](versionoverrides.md) 节点的 URL。清单元素通过使用资源的 **id** 来指定资源。这有助于将清单的大小保持在可管理的范围，尤其是当资源具有不同区域设置的版本时。**id** 在清单内必须是唯一的且最多可包含 32 个字符。

每个资源可以具有一个或多个 **Override** 子元素以定义特定区域设置的不同资源。

## <a name="child-elements"></a>子元素

|  元素 |  类型  |  Description  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  提供指向图标图像的 HTTPS URL。 |
|  **Urls**                |  url     |  提供 HTTPS URL 位置。一个 URL 最多可包含 2048 个字符。 |
|  **ShortStrings** |  字符串  |  **Label** 和 **Title** 元素的文本。每个 **String** 最多可包含 125 个字符。|
|  **LongStrings**  |  string  | **Description** 属性的文本。每个 **String** 最多可包含 250 个字符。|

> [!NOTE]
> 必须对 **Image** 和 **Url** 元素中的所有 URL 使用安全套接字层 (SSL)。

### <a name="images"></a>图像
每个图标必须具有三个**Images**元素，三个强制大小的元素分别为：

- 16x16
- 32x32
- 80x80

此外还支持以下其他大小，但并不是必需的：

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT] 
> Outlook 需要缓存图像资源的能力，以提高性能。 为此，托管图像资源的服务器不能向响应头添加任何 CACHE-CONTROL 指令。 这将导致 Outlook 自动替代泛型或默认图像。    

## <a name="resources-examples"></a>资源示例 

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
