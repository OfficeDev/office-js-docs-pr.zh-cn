---
title: 安装加载项时自动打开任务窗格
description: 了解如何将 Office 加载项配置为在安装时自动打开。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d6ff4b8b5b68236d435ec91b2dcbe121f211081d
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674763"
---
# <a name="automatically-open-a-task-pane-when-an-add-in-is-installed"></a>安装加载项时自动打开任务窗格

可以将外接程序的任务窗格配置为在安装后立即启动。 此功能会增加使用量。 

默认情况下， *不* 包含任何 [加载项命令](../design/add-in-commands.md) 的任务窗格加载项会在安装后立即打开任务窗格。 但是，当加载项具有一个或多个加载项命令时，系统会通知用户新外接程序，但外接程序不会自动启动。 此历史默认行为正在发生变化，因此在某些情况下，具有加载项命令的加载项将自动启动。 此外，如果加载项具有多个任务窗格页，则可以控制加载项是否在安装时启动，如果是，则可以在任务窗格中打开哪个页面。

> [!NOTE]
> 
> - 此功能目前仅在Office web 版中可用。 我们正在努力将此行为引入其他平台，但目前它们仍会显示前面所述的历史默认行为。
> - 此功能仅适用于最终用户安装的加载项，而不适用于集中部署的外接程序。
> - 此功能不适用于内容加载项或邮件 (Outlook) 加载项。
> - 此功能仅适用于具有至少一个 [“任务窗格命令”类型的](../design/add-in-commands.md#types-of-add-in-commands)加载项命令的加载项。

## <a name="new-behavior"></a>新行为

新行为如下所示：

- 如果外接程序只有一个 [任务窗格命令](../design/add-in-commands.md#types-of-add-in-commands)，则选择外接程序的功能区选项卡，并在安装时自动打开任务窗格。 无需配置任何内容。
- 如果外接程序具有多个任务窗格命令，并且其中一个配置为默认 (请参阅) [配置默认任务窗格](#configure-default-task-pane) ，则选择外接程序的功能区选项卡，并在安装时自动打开默认任务窗格。
- 如果外接程序具有多个任务窗格命令，但没有一个配置为默认值，则安装后会自动选择外接程序的功能区选项卡，并在它附近显示标注，通知用户新加载项，但不会打开任何任务窗格。 这与历史默认行为相同。

> [!NOTE]
> 如果出于任何原因，启动任务窗格的外接程序命令在启动时无法由用户手动选择，例如在启动时 [将其配置为禁](../design/disable-add-in-commands.md) 用时，无论配置如何，都不会自动打开它。 

## <a name="configure-default-task-pane"></a>配置默认任务窗格

若要将任务窗格指定为默认值，请将 [TaskpaneId](/javascript/api/manifest/action#taskpaneid) 元素添加为该元素的第一个子 **\<Action\>** 元素，并将其值设置为 **Office.AutoShowTaskpaneWithDocument**。 示例如下。

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

> [!TIP]
> 如果希望加载项在用户重新打开文档时自动启动，则需要执行进一步的配置步骤。 有关何时使用此功能的详细信息和建议，请参阅 [使用文档自动打开任务窗格](automatically-open-a-task-pane-with-a-document.md)。 

## <a name="see-also"></a>另请参阅

- [随文档自动打开任务窗格](automatically-open-a-task-pane-with-a-document.md)
