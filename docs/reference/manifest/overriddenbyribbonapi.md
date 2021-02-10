---
title: 清单文件中 OverriddenByRibbonApi 元素
description: 了解如何指定自定义选项卡、组、控件或菜单项在作为自定义上下文选项卡的一部分时不应显示。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 62aa484057221f9cd7f41af9c8b9210cdb5b3376
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173996"
---
# <a name="overriddenbyribbonapi-element"></a><span data-ttu-id="0b1b9-103">OverriddenByRibbonApi 元素</span><span class="sxs-lookup"><span data-stu-id="0b1b9-103">OverriddenByRibbonApi element</span></span>

<span data-ttu-id="0b1b9-104">指定在支持 API ([](group.md) [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) 的应用程序和平台组合上是否隐藏[CustomTab、](customtab.md)组、按钮控件、菜单控件或菜单项，该 API 将在功能区上安装自定义上下文选项卡。 [](control.md#button-control) [](control.md#menu-dropdown-button-controls)</span><span class="sxs-lookup"><span data-stu-id="0b1b9-104">Specifies whether a [CustomTab](customtab.md), [Group](group.md), [Button](control.md#button-control) control, [Menu](control.md#menu-dropdown-button-controls) control, or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) that installs custom contextual tabs on the ribbon.</span></span>

<span data-ttu-id="0b1b9-105">如果省略它，则默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-105">If it is omitted, the default is `false`.</span></span> <span data-ttu-id="0b1b9-106">如果使用，则它必须是父 *元素* 的第一个子元素。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-106">If it is used, it must be the *first* child element of its parent element.</span></span>

> [!NOTE]
> <span data-ttu-id="0b1b9-107">有关此元素的完全了解，请阅读在不支持自定义上下文选项卡时实现 [备用 UI 体验](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-107">For a full understanding of this element, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

<span data-ttu-id="0b1b9-108">此元素的目的是在外接程序中创建回退体验，该体验在外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时实现自定义上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-108">The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> <span data-ttu-id="0b1b9-109">基本策略是，将某些或所有组和控件从自定义上下文选项卡复制到一个或多个自定义核心选项卡 (即非上下文自定义选项卡) 。 </span><span class="sxs-lookup"><span data-stu-id="0b1b9-109">The essential strategy is that you duplicate some or all of the groups and controls from your custom contextual tab onto one or more custom core tabs (that is, *noncontextual* custom tabs).</span></span> <span data-ttu-id="0b1b9-110">然后，为了确保这些组和控件在不支持自定义上下文选项卡时显示，但在支持自定义上下文选项卡时不显示，可添加为 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab、Group、Control** 或 Menu **Item** 元素的第一个子元素。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-110">Then, to ensure that these groups and controls appear when custom contextual tabs are *not* supported, but do not appear when custom contextual tabs *are* supported, you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **CustomTab**, **Group**, **Control**, or menu **Item** elements.</span></span> <span data-ttu-id="0b1b9-111">这样做的效果如下：</span><span class="sxs-lookup"><span data-stu-id="0b1b9-111">The effect of doing so is the following:</span></span>

- <span data-ttu-id="0b1b9-112">如果加载项在支持自定义上下文选项卡的应用程序和平台上运行，则重复的选项卡、组和控件将不会显示在功能区上。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-112">If the add-in runs on an application and platform that support custom contextual tabs, then the duplicated tabs, groups, and controls won't appear on the ribbon.</span></span> <span data-ttu-id="0b1b9-113">相反，当加载项调用该方法时，将安装自定义上下文 `requestCreateControls` 选项卡。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-113">Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="0b1b9-114">如果加载项在不支持自定义上下文选项卡的应用程序或平台上运行，则重复的选项卡、组和控件将显示在功能区上。</span><span class="sxs-lookup"><span data-stu-id="0b1b9-114">If the add-in runs on an application or platform that *doesn't* support custom contextual tabs, then the duplicated tabs, groups, and controls will appear on the ribbon.</span></span>

## <a name="examples"></a><span data-ttu-id="0b1b9-115">示例</span><span class="sxs-lookup"><span data-stu-id="0b1b9-115">Examples</span></span>

### <a name="overriding-an-entire-tab"></a><span data-ttu-id="0b1b9-116">覆盖整个选项卡</span><span class="sxs-lookup"><span data-stu-id="0b1b9-116">Overriding an entire tab</span></span>

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

### <a name="overriding-a-group"></a><span data-ttu-id="0b1b9-117">替代组</span><span class="sxs-lookup"><span data-stu-id="0b1b9-117">Overriding a group</span></span>

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

### <a name="overriding-a-control"></a><span data-ttu-id="0b1b9-118">替代控件</span><span class="sxs-lookup"><span data-stu-id="0b1b9-118">Overriding a control</span></span>

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

### <a name="overriding-a-menu-item"></a><span data-ttu-id="0b1b9-119">替代菜单项</span><span class="sxs-lookup"><span data-stu-id="0b1b9-119">Overriding a menu item</span></span>


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
