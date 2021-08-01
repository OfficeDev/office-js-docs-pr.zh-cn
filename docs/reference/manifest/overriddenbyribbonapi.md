---
title: 清单文件中 OverriddenByRibbonApi 元素
description: 了解如何指定自定义选项卡、组、控件或菜单项在也是自定义上下文选项卡的一部分时不应显示。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 06d7ae47c7d9de9d48fe9c59493258efc1fa9cae
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671749"
---
# <a name="overriddenbyribbonapi-element"></a>OverriddenByRibbonApi 元素

指定是否在支持 API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)) 的应用程序和平台组合上隐藏[CustomTab、Group、Button](customtab.md)控件、菜单控件或菜单项，该 API 将在功能区上安装自定义上下文选项卡。 [](group.md) [](control.md#button-control) [](control.md#menu-dropdown-button-controls)

如果省略它，则默认值为 `false` 。 如果使用，则它必须是父 *元素* 的第一个子元素。

> [!NOTE]
> 有关此元素的完全了解，请阅读在不支持自定义上下文选项卡时实现 [备用 UI 体验](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

此元素的目的是在外接程序中创建回退体验，当外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时，该外接程序实现自定义上下文选项卡。 基本策略是，将自定义上下文选项卡中的部分或所有组和控件复制到一个或多个自定义核心选项卡上 (自定义选项卡，即自定义选项卡) 。  然后，为了确保这些组和控件在自定义上下文选项卡不受支持时显示，但在支持自定义上下文选项卡时不显示，可添加为 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab、Group、Control** 或 Menu **Item** 元素的第一个子元素。 这样做的效果如下：

- 如果外接程序在支持自定义上下文选项卡的应用程序和平台上运行，则重复的选项卡、组和控件将不会显示在功能区上。 相反，当外接程序调用 方法时，将安装自定义上下文 `requestCreateControls` 选项卡。
- 如果外接程序在不支持自定义上下文选项卡的应用程序或平台上运行，则重复的选项卡、组和控件将显示在功能区上。

## <a name="examples"></a>示例

### <a name="overriding-an-entire-tab"></a>覆盖整个选项卡

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a>替代组

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a>替代控件

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a>替代菜单项


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
