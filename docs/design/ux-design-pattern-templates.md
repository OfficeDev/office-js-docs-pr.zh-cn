# <a name="ux-design-patterns-for-office-add-ins"></a> Office加载项的UX设计模式

Office加载项用户体验设计应该向Office用户提供引人注目的体验，并在默认的Office用户界面中，实现无缝配合，扩展Office整体体验。  

我们的UX模式由各个组件组成。 组件是帮助客户与软件或服务元素进行交互的控件。 按钮、导航和菜单是常见组件的示例，通常具有一致的样式和行为。

Office UI Fabric 呈现外观和行为类似于 Office 部件的组件。 利用Fabric，易于与Office集成。 如果加载项自身预先存在组件语言，则不需要为了Fabric而放弃此语言。 与 Office 集成的同时寻找保留该语言的机会。 寻找置换出风格元素、删除冲突，或采用样式和行为以避免用户混淆的方法。

所提供的模式是基于常见客户场景和用户体验研究的最佳实践解决方案。 它们旨在为设计和开发加载项提供快速切入点，并为实现微软和品牌元素之间的平衡提供指导。 提供明快、现代化的用户体验，在此体验中，来自微软Fabric设计语言的设计元素与合作伙伴独特的品牌标识处于平衡状态，这可能有助于促使用户保留和采用您的加载项。

使用UX模式模板：

* 将解决方案应用于常见的客户场景。
* 应用设计最佳实践。
* 纳入“[Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/get-started)”组件和样式。
* 构建以可视方式与默认 Office 用户界面集成的加载项。
* 想象UX。


## <a name="getting-started"></a>新手指南

这些模式按照加载项中常见的关键操作或体验进行组织。 主要组别是：

* [首次运行体验（FRE）](../design/first-run-experience-patterns.md)
* [身份验证](../design/authentication-patterns.md)
* [导航](../design/navigation-patterns.md)
* [品牌设计](../design/branding-patterns.md)

浏览每个分组，了解如何使用最佳做法设计加载项。



>注：本文档所显示的示例屏幕按照 **1366×768**的分辨率进行设计和显示。




## <a name="see-also"></a>另请参阅
* [设计工具包](design-toolkits.md)
* [Office UI Fabric](https://developer.microsoft.com/en-us/fabric)
* [ Office加载项的最佳开发实践](https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/add-in-development-best-practices)
* [开始使用 Fabric React](https://docs.microsoft.com/en-us/office/dev/add-ins/design/using-office-ui-fabric-react)
