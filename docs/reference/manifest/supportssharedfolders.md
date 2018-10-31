# <a name="supportssharedfolders-element"></a><span data-ttu-id="9cb70-101">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="9cb70-101">SupportsSharedFolders element</span></span>

<span data-ttu-id="9cb70-102">定义 Outlook 加载项是否可在委派方案中可用。</span><span class="sxs-lookup"><span data-stu-id="9cb70-102">Defines whether the Outlook add-in is available in delegate scenarios and is set to false by default.</span></span> <span data-ttu-id="9cb70-103"> *\*SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="9cb70-103">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="9cb70-104">它是默认设置为 *false* 。</span><span class="sxs-lookup"><span data-stu-id="9cb70-104">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9cb70-105">此元素仅在[  Outlook 加载项预览要求设置](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 针对 Exchange 联机中可用。</span><span class="sxs-lookup"><span data-stu-id="9cb70-105">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="9cb70-106">无法将使用此元素的加载项发布到 AppSource 或通过集中部署来部署。</span><span class="sxs-lookup"><span data-stu-id="9cb70-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="9cb70-107">下面是 **SupportsSharedFolders** 元素的示例。</span><span class="sxs-lookup"><span data-stu-id="9cb70-107">The following is an example of the **FunctionFile** element.</span></span>

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
