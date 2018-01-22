# <a name="dropdown-component-in-office-ui-fabric"></a>Office UI Fabric 中的下拉组件

下拉组件是单击下拉按钮后显示的一系列选项。 下拉列表或菜单可用于简化 UI 设计，可方便用户在 UI 内做决定。 列表折叠时，选定项是可见的。 若要更改选定项，用户可以打开下拉列表，并选择一个新值。
  
#### <a name="example-drop-down-in-a-task-pane"></a>示例：任务窗格中的下拉组件

<br/>

![显示下拉组件的图像](../images/overview_withApp_dropdown.png)

<br/>

## <a name="best-practices"></a>最佳做法

|**允许事项**|**不应做**|
|:------------|:--------------|
|如果默认选定选项比其他选项更有可能被选中，应使用下拉组件。 相比之下，ChoiceGroup 或单选按钮显示所有选项，强调了所有选项的同等重要性。|如果所有选项的被选中可能性都相等，不应使用下拉组件。|
|如果有多个选项可以折叠到一个字段中，应使用下拉组件。 另外，如果项列表很长或屏幕空间有限，也应使用下拉组件。|如果选项少于两个，不应使用下拉组件。 应改用复选框。|
|在下拉组件中，应使用缩短的语句或字词。| |

## <a name="variants"></a>变体

|**变体**|**说明**|**示例**|
|:------------|:--------------|:----------|
|**不受控的基本下拉组件**|当多个选项可供选择时使用。|![不受控的基本下拉组件图像](../images/dropdownUncontrolled.png)<br/>|
|**具有 defaultSelectedKey 且已禁用的不受控下拉组件**|状态为已禁用的下拉组件。|![具有 defaultSelectedKey 且已禁用的不受控下拉组件图像](../images/dropdownDisabled.png)<br/>|
|**受控下拉组件**|当默认选定项受 UI 中其他位置影响，且必须保持选定项在下拉组件中的位置时使用。|![受控下拉组件图像](../images/dropdownControlled.png)<br/>|

## <a name="implementation"></a>实现

有关详细信息，请参阅[下拉列表](https://dev.office.com/fabric#/components/dropdown)和 [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)。

## <a name="additional-resources"></a>其他资源

- [用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md)
