---
title: 在 Office 加载项中使用 Office UI Fabric React
description: 了解如何在 Office 加载项中使用 Office UI Fabric React。
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: 7166e9a13c89a1ef2a52659bf31561574f544420
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771335"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>在 Office 加载项中使用 Office UI Fabric React

Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果使用 React 生成外接程序，请考虑使用 Fabric React 来创建用户体验。Fabric 提供了多个可在外接程序中使用的基于 React 的 UX 组件，如按钮或复选框。

本文介绍如何创建使用 React 构建的加载项, 并使用 Fabric React 组件。 

> [!NOTE]
> [Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors)是 Fabric React 附带的，这意味着在完成本文中的步骤后，你的加载项也有权访问 Fabric Core。

## <a name="create-an-add-in-project"></a>创建加载项项目

将使用 Office 加载项的 Yeoman 生成器创建使用 React 的加载项项目。

### <a name="install-the-prerequisites"></a>安装必备组件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a>创建项目

使用 Yeoman 生成器创建 Word 加载项项目。 运行下面的命令，再回答如下所示的提示问题：

```command&nbsp;line
yo office
```

- **选择项目类型:** `Office Add-in Task Pane project using React framework`
- **选择脚本类型:** `TypeScript`
- **要如何命名加载项?** `My Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** `Word`

![Yeoman 生成器](../images/yo-office-word-react.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

### <a name="try-it-out"></a>试用

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. 完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。

    > [!NOTE]
    > Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

    > [!TIP]
    > 如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。 运行此命令时，本地 Web 服务器将启动。
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - 若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。 这将启动本地的 Web 服务器 (如果尚未运行的话), 并使用加载的加载项打开 Word。

        ```command&nbsp;line
        npm start
        ```

    - 若要在浏览器版 Word 中测试加载项，请在项目的根目录中运行以下命令。 如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。

        ```command&nbsp;line
        npm run start:web
        ```

        若要使用加载项，请在 Word 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。

3. 在 Word 中，依次选择“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。 请注意任务窗格底部的“默认文本”和 "**运行**" 按钮。 在本演练的其余部分中, 你将通过创建使用来自 Fabric React 的 UX 组件的 React 组件来重新定义此文本和按钮。

    ![Word 应用程序的屏幕截图，任务窗格中突出显示了 "显示任务窗格" 功能区按钮以及“运行……”按钮和前面的文本](../images/word-task-pane-yo-default.png)


## <a name="create-a-react-component-that-uses-fabric-react"></a>创建使用 Fabric React 的 React 组件

此时, 你已经创建了一个使用 React 构建的非常基本的任务窗格加载项。 接下来，完成以下步骤，在加载项项目中创建新的 React 组件 (`ButtonPrimaryExample`)。 该组件使用 Fabric React 的`Label`和`PrimaryButton`组件。

1. 打开 Yeoman 生成器创建的项目文件夹，并转到**src\taskpane\components**。
2. 在该文件夹中，创建一个名为“**Button.tsx**”的新文件。
3. 在 **Button.tsx** 中，输入以下代码以定义`ButtonPrimaryExample`组件。

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // In the click event, write text to the document.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
      await context.sync();
    });
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

此代码将执行以下操作：

- 引用使用 `import * as React from 'react';` 的 React 库。
- 参考用于创建 `ButtonPrimaryExample` 的 Fabric 组件（`PrimaryButton`、`IButtonProps`、`Label`）。
- 声明新的 `ButtonPrimaryExample` 组件使用 `export class ButtonPrimaryExample extends React.Component`。
- 声明 `insertText` 将处理按钮 `onClick` 事件的函数。
- 定义 `render` 函数中 React 组件的 UI。 HTML 标记使用 Fabric Reac 中的组件 `Label` 和 `PrimaryButton`，并指定当 `onClick` 事件触发时，`insertText` 函数将运行。

## <a name="add-the-react-component-to-your-add-in"></a>将 React 组件添加到加载项

通过打开 **src\components\App.tsx** 并完成下列步骤，将组件 `ButtonPrimaryExample` 添加到加载项：

1. 添加下面导入语句，以`ButtonPrimaryExample`从**Button.tsx**中引用。

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. 删除以下两个导入语句。

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. 将默认 `render()` 函数替换为以下使用 `ButtonPrimaryExample` 的代码。

    ```typescript
    render() {
      return (
        <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
          <ButtonPrimaryExample />
        </HeroList>
        </div>
      );
    }
    ```

  4. 将所做的更改保存到**App.tsx**。

## <a name="see-the-result"></a>查看结果

在 Word 中, 当你保存对**App.tsx**的更改时，加载项任务窗格会自动更新。 任务窗格底部的默认文本和按钮现在显示由该`ButtonPrimaryExample`组件定义的 UI。 选择**插入文本……** 按钮将文本插入到文档中。

![Word 应用程序的屏幕截图，突出显示 "插入文本……" 按钮和前面的文本](../images/word-task-pane-with-react-component.png)

恭喜，您已使用 React 和 Office UI Fabric React 成功创建了任务窗格加载项！ 

## <a name="see-also"></a>另请参阅

- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md)
- [Office UI Fabric React](https://developer.microsoft.com/fabric)
- [适用于 Office 加载项的 UX 设计模式](ux-design-pattern-templates.md)
- [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
