---
title: 有关在政府云上部署 Office 加载项的指南
description: 了解如何部署 Office 加载项来保护政府云环境
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: f3995c62a1b7fb482df6a15da870f747f55e9508
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607589"
---
# <a name="guidance-for-deploying-office-add-ins-on-government-clouds"></a>有关在政府云上部署 Office 加载项的指南

Microsoft 为本地、州和国家政府组织中隐私敏感的客户提供政府云选项。 这为合作伙伴提供了使用其 Office 外接程序面向关键客户的机会。由于这些环境的性质受到限制，这对客户的隐私和安全需求非常重要，因此，并非标准生产环境中通常提供的所有资源都可在这些云中使用。

对于在这些受限制的云环境中向客户提供 Office 加载项的合作伙伴，必须考虑与标准公有云环境的重要区别。 以下信息提供了需要开发人员编写面向这些环境中客户的 Office 加载项的特殊处理的详细信息。

## <a name="all-sovereign-environments"></a>所有主权环境

对于所有政府云 (（即主权云) 环境），公共 Office 应用商店不可用。 这意味着最终用户无法直接从公共商店获取 Office 加载项。 管理员也无法直接从公共商店将 Office 加载项部署到其管理员门户。 相反，必须与管理员合作，以确保以下各项：

- 解决方案所需的资源和服务在云边界内可用。 可以与租户管理员合作，在云边界内预配服务和资源，或者与网络管理员合作，以启用对驻留在云边界之外的资源的访问。

- Office 外接程序访问的资源符合要从中访问它们的政府云的要求。 它们还必须符合为其预配解决方案的客户租户的任何其他要求。 这些要求包括潜在敏感数据的传输、管理和存储，以及代码和资源安全性以及访问审查过程。

- 介绍解决方案及其适用于特定政府云部署的源位置的 Office 外接程序清单是从合作伙伴获取的，并通过管理员门户引入以部署到相应的用户。

## <a name="us-government-community-cloud-gcc"></a>美国政府社区云 (GCC) 

除了适用于所有主权云的要求外，每个主权云环境都有自己的注意事项，可能会影响面向环境的 Office 加载项。 GCC 是政府云环境限制最少的，外接程序所需的一些资源可从此环境中的常规生产终结点获得。 其中一个资源是 Office JavaScript API 库。 解决方案合作伙伴可以继续引用公共Office.js资源，就像使用公共生产解决方案一样。

## <a name="gcc-high-gcch-us-department-of-defense-dod-or-other-sovereign-managed-clouds"></a>GCC 高 (GCCH) 、美国国防部 (DOD) 或其他主权托管云

这些政府云仍与 Internet 连接，但提供的公共终结点集受到严格限制。 其中一个受限终结点是用于加载 Office JavaScript API 库的公共终结点。 无法从这些环境中访问Office.js的公共 CDN 位置。 但是，将预配具有相同资源的内部、每云 Microsoft Office CDN。 这意味着用于访问Office.js的终结点 URL 将有所不同，Office 加载项可能需要运行某种级别的自定义。 鉴于其他限制，提供给客户的任何解决方案都可能需要在环境中托管提供程序服务，因此需要其他自定义项。 你需要确定向客户提供解决方案的最佳方式，以便它符合对在这些环境边界内运行的服务施加的其他限制。

## <a name="airgapped-sovereign-clouds"></a>空封的主权云

这些政府云基本上与公共互联网完全断开连接。 通常从公共资源访问的任何资源都必须在这些云环境中进行自定义预配。 在前面提到的 GCCH 和 DOD 云中，大多数 (（如果不是所有) 解决方案提供商都需要在云中预配其服务和资源）。 可以选择创建允许访问公共服务和资源的防火墙异常。 但是，无法在映射的云中绕过此方法。 与 GCCH 和 DOD 云一样，每个云环境中都会预配一个 Microsoft Office CDN，用于托管所需的资源，例如Office.js库。 你需要与客户租户管理员密切合作，确定以符合已封陆主权云的严格访问要求的方式提供服务和资源的最佳方式。
