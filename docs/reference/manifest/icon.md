---
title: 清单文件中的 Icon 元素
description: 定义“按钮”或“菜单”控件的 Image 元素。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9eb4ccf394bb1c894f2b17f34038ca64fee09dc5
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341063"
---
# <a name="icon-element"></a>图标元素

为"按钮"或 **"菜单** " [控件定义一组](control-button.md) [Image](control-menu.md) 元素。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  否  | 正在定义的图标类型。这仅适用于移动外形规格中的图标。[MobileFormFactor](mobileformfactor.md) 元素中所包含的 **Icon** 元素必须将此属性设置为 `bt:MobileIconList`。 |

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Image](#image)        | 是 |   要使用的图像的 resid         |

### <a name="image"></a>图像

按钮的图像。 **resid** 属性不能超过 32 个字符，必须设置为 **Images** 元素（位于 [Resources](resources.md) 元素）中 **Image** 元素的 **id** 属性的值。 The **size** attribute indicates the size in pixels of the image. 有三个图像大小为必需（16、32 和 80 像素），另外还支持五个大小（20、24、40、48 和 64 像素）。

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> 如果此图像是加载项的代表图标，请参阅在 [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) 和加载项内创建有效Office大小和其他要求。

## <a name="additional-requirements-for-mobile-form-factors"></a>移动外形规格的其他要求

当父 **Icon** 元素是 [MobileFormFactor](mobileformfactor.md) 元素的后代时，所要求的最小大小会略有不同。 清单必须至少提供 25、32 和 48 像素大小。 所提供的每个大小必须出现三次，并将 `scale` 属性设置为 `1`、`2` 或 `3`。 此属性指定 `UIScreen.scale` iOS 设备的属性。 有关详细信息，请参阅 [scale](https://developer.apple.com/documentation/uikit/uiscreen/1617836-scale)。

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
