# <a name="group-element"></a>Group 元素

在选项卡中定义 UI 控件组。在自定义选项卡上，加载项最多可以创建 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。加载项限于一个自定义选项卡。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  是  | 组的唯一 ID。|

### <a name="id-attribute"></a>id 属性

必需。组的唯一标识符。是一个最多为 125 个字符的字符串。该字符串在清单内必须是唯一的，否则组将无法呈现。

## <a name="child-elements"></a>子元素
|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)      | 是 |  CustomTab 或组的标签。  |
|  [Control](#control)    | 是 |  一个或多个 Control 对象的集合。  |

### <a name="label"></a>标签 

必需。组的标签。**resid** 属性必须设置为 [Resources](resources.md) 元素的 **ShortStrings** 元素中的 **String** 元素的 **id** 属性的值。

### <a name="control"></a>控件
一个组至少需要一个控件。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```