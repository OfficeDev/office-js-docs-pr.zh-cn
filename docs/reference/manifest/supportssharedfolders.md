# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 元素

定义 Outlook 加载项是否可在委派方案中可用。  **SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md)的子元素。 它是默认设置为 *false* 。

> [!IMPORTANT]
> 此元素仅在[  Outlook 加载项预览要求设置](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 针对 Exchange 联机中可用。 无法将使用此元素的加载项发布到 AppSource 或通过集中部署来部署。

下面是 **SupportsSharedFolders** 元素的示例。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
