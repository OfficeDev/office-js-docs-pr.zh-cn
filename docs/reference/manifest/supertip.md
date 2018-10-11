# <a name="supertip"></a>Supertip

定义丰富的工具提示（标题和说明）。它由[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件使用。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  描述  |
|:-----|:-----|:-----|
|  [标题](#title)        | 是 |   Supertip 的文本。         |
|  [描述](#description)  | 是 |  Supertip 的说明。    |

### <a name="title"></a>标题

必需。SuperTip 的文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。

### <a name="description"></a>描述

必需。SuperTip 的描述。**resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。

## <a name="example"></a>示例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
