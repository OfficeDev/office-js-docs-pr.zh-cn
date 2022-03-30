---
title: 部署和安装 Outlook 加载项以进行测试
description: 创建清单文件，将加载项 UI 文件部署到 Web 服务器，在邮箱中安装加载项，然后测试加载项。
ms.date: 07/08/2021
ms.localizationpriority: high
ms.openlocfilehash: b627dbf4b32daee4327cb139db58a56c4a704580
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496878"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>部署和安装 Outlook 加载项以进行测试

在开发 Outlook 加载项的过程中，可能会发现自己在反复部署和安装加载项以进行测试，会涉及以下步骤。

1. 创建描述外接程序的清单文件。
1. 将外接程序 UI 文件部署到 Web 服务器。
1. 在邮箱中安装外接程序。
1. 测试加载项，对 UI 或清单文件进行适当更改，并重复步骤 2 和步骤 3 来测试更改。

> [!NOTE]
> [已弃用自定义窗格](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)，因此请确保正在使用[受支持的加载项扩展点](outlook-add-ins-overview.md#extension-points)。

## <a name="create-a-manifest-file-for-the-add-in"></a>创建加载项清单文件

每个外接程序都通过一个 XML 清单进行描述，该文档为服务器提供有关外接程序的信息，为用户提供外接程序的描述性信息，并标识外接程序 UI HTML 文件的位置。您可以在本地文件夹或服务器上存储该清单，只要所测试的邮箱的 Exchange 服务器能够访问这个位置即可。我们假定您在本地文件夹中存储清单。有关如何创建清单文件的信息，请参阅 [Outlook 外接程序清单](manifests.md)。

## <a name="deploy-an-add-in-to-a-web-server"></a>将加载项部署到 Web 服务器

可以使用 HTML 和 JavaScript 创建外接程序。生成的源文件存储在 Web 服务器上，可供托管外接程序的 Exchange 服务器进行访问。在最初部署外接程序的源文件后，可以将 Web 服务器上存储的 HTML 文件或 JavaScript 文件替换为 HTML 文件的新版本，从而更新外接程序 UI 和行为。

## <a name="install-the-add-in"></a>安装加载项

准备好外接程序清单文件，并将外接程序 UI 部署到可供访问的 Web 服务器后，可以使用 Outlook 客户端为 Exchange 服务器上的邮箱旁加载外接程序，也可以通过运行远程 Windows PowerShell cmdlet 安装外接程序。

### <a name="sideload-the-add-in"></a>旁加载加载项

如果邮箱位于 Exchange Online、Exchange 2013 或更高版本上，可以安装外接程序。至少必须拥有 Exchange Server 的 **我的自定义应用程序** 角色，才能旁加载外接程序。若要测试外接程序，或通过指定外接程序清单的 URL 或文件名来常规安装外接程序，应让 Exchange 管理员提供必要权限。

Exchange 管理员可以运行下列 PowerShell cmdlet，向一个用户分配必要权限。在下面的示例中，`wendyri` 是用户的电子邮件别名。

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

如有必要，管理员可以运行下列 cmdlet，向多个用户分配类似的必要权限。

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

有关我的自定义应用角色的详细信息，请参阅[我的自定义应用角色](/exchange/my-custom-apps-role-exchange-2013-help)。

在使用 Microsoft 365 或 Visual Studio 开发加载项时，会向你分配组织管理员角色，这便允许你按 EAC 中的文件或 URL 或者按 Powershell cmdlet 安装加载项。

### <a name="install-an-add-in-by-using-remote-powershell"></a>使用远程 PowerShell 安装加载项

在 Exchange 服务器上创建远程 Windows PowerShell 会话后，可以运行 `New-App` cmdlet 和下列 PowerShell 命令，安装 Outlook 外接程序。

```powershell
New-App -URL:"http://<fully-qualified URL">
```

完全限定的 URL 是为外接程序准备的外接程序清单文件的位置。

使用下列附加 PowerShell cmdlet，管理邮箱的加载项。

- `Get-App` - 列出为邮箱启用的外接程序。
- `Set-App` - 在邮箱中启用或禁用外接程序。
- `Remove-App` - 从 Exchange 服务器中删除以前安装的外接程序。

## <a name="client-versions"></a>客户端版本

若要决定测试什么版本的 Outlook 客户端，请综合考虑自己的开发需求。

- 若要开发供私人使用或仅供组织成员使用的外接程序，请务必测试公司使用的 Outlook 版本。请注意，某些用户可能会使用 Outlook 网页版。因此，还请务必测试公司的标准浏览器版本。

- 如果开发的是要在 [AppSource](https://appsource.microsoft.com) 中列出的加载项，必须测试 [商业市场认证策略 1120.3](/legal/marketplace/certification-policies#11203-functionality) 中指定的必需版本。这包括:
  - 最新版 Windows 版 Outlook 和前一个版本。
  - 最新版 Mac 版 Outlook。
  - 最新 iOS 版和 Android 版 Outlook（如果加载项[支持移动设备规格](add-mobile-support.md)）。
  - 商业市场验证策略 1120.3 中指定的浏览器版本。

> [!NOTE]
> 如果由于[请求的 API 要求集](apis.md)不受客户端支持，导致外接程序不支持上述客户端之一，将从所需客户端列表中删除相应客户端。

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Outlook 网页版和 Exchange 服务器版本

在访问 Outlook 网页版时，消费者和 Microsoft 365 帐户用户将看到新式 UI 版本，而不会再看到已弃用的经典版本。 但是，本地 Exchange 服务器将继续支持经典 Outlook 网页版。 因此，在验证过程中，你的提交可能会收到一条警告，指出加载项与经典 Outlook 网页版不兼容。 在这种情况下，需考虑在本地 Exchange 环境中测试加载项。 此警告不会阻止你向 AppSource 提交加载项，但如果消费者是在本地 Exchange 环境中使用 Outlook 网页版，则可能无法获得最佳的体验。

为缓解此问题，我们建议你在连接到你自己的专用本地 Exchange 环境的 Outlook 网页版中对加载项进行测试。 有关详细信息，请参阅有关如何[建立 Exchange 2016 或 Exchange 2019 测试环境](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment)的指南以及有关如何管理[Exchange 服务器中的 Outlook 网页版](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true)的指南。

或者，你也可以选择付费并使用托管和管理本地 Exchange 服务器的服务。有两个选项：

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/microsoft-exchange/)

此外，如果不想面向连接到本地 Exchange 的用户提供自己的加载项，可将加载项清单中的[要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#exchange-server-support)设置为 1.6 或更高版本。 在经典 Outlook 网页版上，不会对此类加载项进行测试或验证。

## <a name="see-also"></a>另请参阅

- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)
