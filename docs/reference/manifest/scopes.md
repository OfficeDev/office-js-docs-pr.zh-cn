# <a name="scopes-element"></a><span data-ttu-id="a8994-101">Scopes 元素</span><span class="sxs-lookup"><span data-stu-id="a8994-101">Scopes element</span></span>

<span data-ttu-id="a8994-102">包含外接程序需要拥有的 Microsoft Graph 访问权限。</span><span class="sxs-lookup"><span data-stu-id="a8994-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="a8994-103">Office 应用商店使用 Scopes 元素创建许可对话框。</span><span class="sxs-lookup"><span data-stu-id="a8994-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="a8994-104">当用户安装应用商店中的外接程序时，系统会提示他们授予外接程序对用户 Microsoft Graph 数据的指定访问权限。</span><span class="sxs-lookup"><span data-stu-id="a8994-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a8994-105">子元素</span><span class="sxs-lookup"><span data-stu-id="a8994-105">Child elements</span></span>

|  <span data-ttu-id="a8994-106">元素</span><span class="sxs-lookup"><span data-stu-id="a8994-106">Element</span></span> |  <span data-ttu-id="a8994-107">类型</span><span class="sxs-lookup"><span data-stu-id="a8994-107">Type</span></span>  |  <span data-ttu-id="a8994-108">说明</span><span class="sxs-lookup"><span data-stu-id="a8994-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a8994-109">**范围**</span><span class="sxs-lookup"><span data-stu-id="a8994-109">**Scope**</span></span>                |  <span data-ttu-id="a8994-110">字符串</span><span class="sxs-lookup"><span data-stu-id="a8994-110">string</span></span>     |   <span data-ttu-id="a8994-111">Microsoft Graph 权限的名称，例如，Files.Read.All。</span><span class="sxs-lookup"><span data-stu-id="a8994-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="a8994-112">示例</span><span class="sxs-lookup"><span data-stu-id="a8994-112">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
