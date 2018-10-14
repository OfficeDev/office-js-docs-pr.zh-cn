# <a name="desktopformfactor-element"></a>DesktopFormFactor 元素

指定对桌面外形规格的加载项设置。桌面外形规格包括 Office for Windows、Office for Mac 和 Office Online。它包含桌面外形规格的所有加载项信息（**Resource** 节点除外）。

每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

> [!IMPORTANT]
> SupportsSharedFolders 元素仅在针对 Exchange Online 的 Outlook 加载项预览要求集中可用。
> 使用此元素的加载项无法出现在 Office 商店或集中部署中。

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | 说明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | 是      | 定义加载项公开功能的位置。 |
| [FunctionFile](functionfile.md)       | 是      | 包含 JavaScript 函数的文件的 URL。|
| [GetStarted](getstarted.md)           | 否       | 定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。 |
| SupportsSharedFolders                 | 否       | 定义 Outlook 加载项在委派方案中是否可用，默认设置为 *false* 。 预览要求集|

## <a name="desktopformfactor-example"></a>DesktopFormFactor 示例

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
