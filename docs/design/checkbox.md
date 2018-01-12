# <a name="checkbox-component-in-office-ui-fabric"></a>Office UI Fabric 中的复选框组件

复选框是一种 UI 元素，可方便用户在加载项内选中或清除选项。复选框可方便用户在选项之间进行选择。 另外，复选框可以与相关控件配对。 选中或清除复选框后，相关控件的行为也会随之变化。 例如，相关控件可能会在可见或隐藏状态之间切换。
  
#### <a name="example-check-box-in-a-task-pane"></a>示例：任务窗格中的复选框

<br/>

![显示复选框的图像](../../images/overview_withApp_checkbox.png)

<br/>

## <a name="best-practices"></a>最佳做法

|**允许事项**|**不应做**|
|:------------|:--------------|
|应使用复选框指明状态。<br/><br/>![“应做”复选框示例](../../images/checkboxDo.png)<br/>|不应使用复选框显示/指明操作。<br/><br/>![“不应做”复选框示例](../../images/checkboxDont.png)<br/>|
|如果用户可以选择多个选项且选项不互斥，应使用多个复选框。|如果用户只能选择一个选项，不应使用复选框。 只需要选择一个选项时，请使用单选按钮。|
|当多个复选框组合到一起时，应支持用户选择任意组合的选项。|不应让两组复选框彼此相邻。 请用标签分隔两组复选框。|
|应对辅助设置使用一个复选框。 例如，“记住我?”****复选框是登录方案中使用的辅助设置。|不应使用复选框启用或禁用设置。 若要更改启用或禁用状态，请使用切换组件。|

## <a name="variants"></a>变体

|**变体**|**说明**|**示例**|
|:------------|:--------------|:----------|
|**不受控复选框**|用作默认复选框状态。 |![不受控复选框图像](../../images/checkbox_unchecked.png)|
|**默认选中的不受控复选框**|当复选框示例保持自身状态时使用。 |![默认选中的不受控复选框图像](../../images/checkbox_checked.png)|
|**默认选中且已禁用的不受控复选框**|状态为已禁用的复选框。 |![默认选中且已禁用的不受控复选框图像](../../images/checkbox_disabled.png)|
|**受控复选框**|是否选中此复选框由 UI 中的其他位置决定。 在此方案中，正确的值会通过 **onChange** 事件和重新呈现 UI 传递给复选框。 |![受控复选框图像](../../images/checkbox_unchecked.png)|

## <a name="implementation"></a>实现

有关详细信息，请参阅[复选框](https://dev.office.com/fabric#/components/checkbox)和 [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)。

## <a name="additional-resources"></a>其他资源

- [用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md)
