# <a name="toggle-component-in-office-ui-fabric"></a>Office UI Fabric 中的切换组件

切换组件表示一种可实现开/关功能的物理开关。 切换组件可用于呈现两个互斥选项（例如，开/关），用户选择任一选项，系统都会立即执行操作。
  
#### <a name="example-toggle-in-a-task-pane"></a>示例：任务窗格中的切换


![显示切换组件的图像](../../images/overview_withApp_toggle.png)

<br/>

## <a name="best-practices"></a>最佳做法

|**允许事项**|**不应做**|
|:------------|:--------------|
|如果更改会立即应用，应对二项设置使用切换组件。<br/><br/>![“应做”切换组件示例](../../images/toggleDo.png)<br/>|如果用户必须执行额外步骤，更改才能生效，不应使用切换组件。<br/><br/>![“不应做”切换组件示例](../../images/toggleDont.png)<br/>|
|只有当要对设置使用多个特定标签时，才应替换“开”****和“关”****标签。 应使用短（3-4 个字符）标签表示相反的选项。| |

## <a name="variants"></a>变量

|**变体**|**说明**|**示例**|
|:------------|:--------------|:----------|
|**已启用并选中**|当切换的状态处于活动时使用。|![已启用并选中图像](../../images/toggleEnabledOn.png)<br/>|
|**已启用但取消选中**|当切换的状态处于非活动时使用。|![已启用但取消选中图像](../../images/toggleEnabledOff.png)<br/>|
|**已禁用并选中**|不能更改活动状态时使用。|![已禁用并选中图像](../../images/toggleDisabledOn.png)<br/>|
|**已禁用但取消选中**|无法更改非活动状态时使用。|![已禁用但取消选中图像](../../images/toggleDisabledOff.png)<br/>|

## <a name="implementation"></a>实现

有关详细信息，请参阅[切换](https://dev.office.com/fabric#/components/toggle)和 [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)。

## <a name="additional-resources"></a>其他资源

- [用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md)
