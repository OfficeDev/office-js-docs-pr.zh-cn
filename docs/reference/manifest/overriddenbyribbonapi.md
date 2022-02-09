---
title: 清单文件中 OverriddenByRibbonApi 元素
description: 了解如何指定自定义选项卡、组、控件或菜单项在也是自定义上下文选项卡的一部分时不应显示。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48977691ee4bf2ccd71bc146647dae452ce9e2fc
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467685"
---
# <a name="overriddenbyribbonapi-element"></a>OverriddenByRibbonApi 元素

指定是否在支持在[](group.md)功能区上安装自定义[](control-menu.md)上下文选项卡的 API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1))) 的应用程序和平台组合上隐藏组、按钮控件、菜单控件或菜单项。 [](control-button.md)

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [功能区 1.2](../requirement-sets/add-in-commands-requirement-sets.md) (Word.Excel、PowerPoint 和 Word.) 

如果省略此元素，则默认值为 `false`。 如果已使用，则它必须是父 *元素* 的第一个子元素。

> [!NOTE]
> 有关此元素的完全了解，请阅读在不支持自定义上下文选项卡时实现 [备用 UI 体验](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

此元素的目的是在外接程序中创建回退体验，当外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时，该外接程序实现自定义上下文选项卡。 基本策略是，将自定义上下文选项卡中的某些或所有组和控件复制到一个或多个自定义核心选项卡上 (即非上下文自定义选项卡) 。  `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`然后，为了确保这些组和控件在自定义上下文选项卡不受支持时显示，但在支持自定义上下文选项卡时不显示，可添加为 **Group**、**Control** 或 menu **Item** 元素的第一个子元素。 这样做的效果如下：

- 如果外接程序在支持自定义上下文选项卡的应用程序和平台上运行，则重复的组和控件将不会显示在功能区上。 相反，当外接程序调用 方法时，将安装自定义上下文 `requestCreateControls` 选项卡。
- 如果外接程序在不支持自定义上下文选项卡的应用程序或平台上运行，则重复的组和控件将显示在功能区上。

## <a name="examples"></a>示例

### <a name="overriding-a-group"></a>替代组

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.CustomTab1.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="Contoso.MyButton1">
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
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.CustomTab2.group2">
      <Control  xsi:type="Button" id="Contoso.MyButton2">
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
  <CustomTab id="Contoso.TabCustom3">
    <Group id="Contoso.CustomTab3.group3">
      <Control  xsi:type="Menu" id="Contoso.MyMenu">
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
