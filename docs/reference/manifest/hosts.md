---
title: 清单文件中的 Hosts 元素
description: 指定Office外接程序将Office的客户端应用程序。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9ea6cc9745f47b6e9b1c9bb0232b744304078053
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341070"
---
# <a name="hosts-element"></a>Hosts 元素

指定Office外接程序将Office的客户端应用程序。 包含 **Host** 元素及其设置的集合。 

## <a name="as-child-of-versionoverrides-element"></a>作为 VersionOverrides 元素的子元素

本节中的信息仅适用于 **Hosts** 元素是 [VersionOverrides 的子级的情况](versionoverrides.md)。

此元素替代基本 **清单中的 Hosts** 元素。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  是   |  说明主机及其设置。 |
