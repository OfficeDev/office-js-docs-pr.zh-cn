# <a name="textfield-component-in-office-ui-fabric"></a>Office UI Fabric 中的 TextField 组件

文本字段可方便用户键入文本。 它通常用于捕获单行文本，但也可以配置为捕获多行文本。 此文本以简单、统一的格式显示在屏幕上。
  
#### <a name="example-textfield-in-a-task-pane"></a>示例：任务窗格中的 TextField

![显示 TextField 的图像](../../images/overview_withApp_textField.png)

<br/>

## <a name="best-practices"></a>最佳做法

|**允许事项**|**不应做**|
|:------------|:--------------|
|应使用文本字段接受表单或页面上的数据输入。|不应使用文本字段在页面的正文元素中呈现基本副本。|
|应用实用名称标记文本字段。|不应将文本字段用于输入日期或时间。 请改用日期/时间选取器。|
|应使用简洁的占位符文本，指定应输入哪些内容。|如果可以预定义有效的输入选项，不应使用文本字段。 请改用下拉组件。|
|应提供文本字段的所有相应状态（静态、悬停、焦点、互动、不可用、错误）。||
|应明确标记必填和可选文本字段。||
|应尽可能根据预期数据格式设置文本字段的格式。 例如，如果捕获 10 位数电话号码，请使用三个独立字段来存储电话号码的不同部分。||

## <a name="variants"></a>变体

|**变体**|**说明**|**示例**|
|:------------|:--------------|:----------|
|**默认 TextField**|用作默认的文本字段。|![默认 TextField 图像](../../images/textfieldDefault.png)<br/>|
|**已禁用 TextField**|当文本字段处于禁用状态时使用。|![已禁用 TextField 图像](../../images/textfieldDisabled.png)<br/>|
|**必填 TextField**|当必须在文本字段中输入内容时使用。|![必填 TextField 图像](../../images/textfieldRequired.png)<br/>|
|**带有占位符的 TextField**|当需要占位符文本时使用。|![带有占位符的 TextField 图像](../../images/textfieldPlaceholder.png)<br/>|
|**多行的 TextField**|当需要多行文本时使用。|![带有占位符的 TextField 图像](../../images/textfieldMulti.png)<br/>|

## <a name="implementation"></a>实现

有关详细信息，请参阅 [TextField](https://dev.office.com/fabric#/components/textfield) 和 [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)。

## <a name="additional-resources"></a>其他资源

- [用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md)
