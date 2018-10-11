# <a name="script-element"></a>Script 元素

定义 Excel 中的自定义函数所使用的脚本设置。

## <a name="attributes"></a>属性

None

## <a name="child-elements"></a>子元素

|元素  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 自定义函数所使用的 JavaScript 文件的资源 id 的字符串。|

## <a name="example"></a>示例

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
