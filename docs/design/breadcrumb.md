---
title: Office UI Fabric 中的痕迹导航组件
description: ''
ms.date: 12/04/2017
---



# <a name="breadcrumb-component-in-office-ui-fabric"></a>Office UI Fabric 中的痕迹导航组件

在 Office 加载项中，痕迹导航组件可用于导航。它们会显示当前页在层次结构中的位置，并帮助用户了解他们相对于层次结构其余部分的位置。此外，痕迹导航组件还支持一键导航到层次结构中的更高级别。
  
#### <a name="example-breadcrumb-in-a-task-pane"></a>示例：任务窗格中的痕迹导航组件

![显示痕迹导航组件的图像](../images/overview-with-app-breadcrumb.png)

## <a name="best-practices"></a>最佳做法

|**允许事项**|**不应做**|
|:------------|:--------------|
|应将痕迹导航组件置于加载项布局顶部、位于项列表之上或位于布局的主内容之上。<br/><br/>![“应做”痕迹导航组件图像](../images/breadcrumb-do.png) |不应将痕迹导航组件用作转到其他页面的主要方法。<br/><br/>![“不应做”痕迹导航组件图像](../images/breadcrumb-dont.png)|

## <a name="implementation"></a>实现

有关详细信息，请参阅[痕迹导航](https://dev.office.com/fabric#/components/breadcrumb)和 [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)。

## <a name="see-also"></a>另请参阅

- [用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md)
