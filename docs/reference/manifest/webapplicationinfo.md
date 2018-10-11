# <a name="webapplicationinfo-element"></a>WebApplicationInfo 元素

支持 Office 加载项中的单一登录 (SSO)。此元素包含加载项中的信息，如下所示：

- OAuth 2.0 *资源*，Office 主机应用程序可能需要访问该资源的权限。
- OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。

**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。  

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **ID**    |  是   |  在 Azure Active Directory v2.0 端点中注册的加载项关联服务的**应用程序 ID**。|
|  **资源**  |  是   |  指定在 Azure Active Directory v2.0 端点中注册的加载项的**应用程序 ID URI**。|
|  [范围](scopes.md)                |  否  |  指定加载项需要拥有的对 Microsoft Graph 的访问权限。  |

> [!NOTE] 
> 目前，加载项的 Resource 必须与其 Host 一致。 Office 不会请求获取加载项令牌，除非可以证明所有权。目前，这是通过在 Resource 的完全限定的域名下托管加载项来完成。

## <a name="webapplicationinfo-example"></a>WebApplicationInfo 示例

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
