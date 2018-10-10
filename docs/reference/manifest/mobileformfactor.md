# <a name="mobileformfactor-element"></a>MobileFormFactor 元素

指定对移动外形规格的外接程序的设置。它包含移动外形规格的所有外接程序信息（**资源**节点的信息除外）。

每个 **MobileFormFactor** 定义均包含 **FunctionFile** 元素以及一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

在 VersionOverrides 架构 1.1 中定义了 **MobileFormFactor** 元素。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | 说明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](functionfile.md)     | 是      | 包含 JavaScript 函数的文件的 URL。|

## <a name="mobileformfactor-example"></a>MobileFormFactor 示例

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
