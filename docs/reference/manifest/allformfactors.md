# <a name="allformfactors-element"></a>AllFormFactors 元素

指定加载项的所有外观设置。 目前，使用  **AllFormFactors** 的唯一功能是自定义函数。 **AllFormFactors** 是使用自定义函数时的必要元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  是 |  定义外接程序公开功能的位置。 |

## <a name="allformfactors-example"></a>AllFormFactors 示例

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
