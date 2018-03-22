---
title: 在 Office 加载项中使用 Office UI Fabric React
description: ''
ms.date: 12/04/2017
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>在 Office 加载项中使用 Office UI Fabric React

Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果使用 React 生成外接程序，请考虑使用 Fabric React 来创建用户体验。Fabric 提供了多个可在外接程序中使用的基于 React 的 UX 组件，如按钮或复选框。

若要开始在加载项中使用 Fabric React 组件，请执行以下步骤。

> [!NOTE]
> 如果按照本文中的步骤操作，也可以在加载项中使用 Fabric Core。

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a>第 1 步 - 使用适用于 Office 的 Yeoman 生成器创建项目

若要创建使用 Fabric React 的外接程序，我们建议使用 Office 的 Yeoman 生成器。Office 的 Yeoman 生成器提供开发 Office 外接程序所需的项目基架和版本管理。

若要创建项目，请使用 **Windows PowerShell**（而不是命令提示符）执行以下步骤：

1. 安装必备组件。
2. 运行 `yo office`，为外接程序创建项目文件。
3. 当系统提示你选择一个 Office 客户端应用程序时，请选择 **Word**。
4. 确保位于包含项目文件的目录中，再运行 `npm start`。此时，显示旋转图标的浏览器窗口自动打开。
5. [旁加载清单](..\testing\test-debug-office-add-ins.md)，以查看加载项的完整 UI。

## <a name="step-2---add-a-fabric-react-component"></a>第 2 步 - 添加 Fabric React 组件

接下来，将 Fabric React 组件添加到外接程序。创建称为 `ButtonPrimaryExample` 的新的 React 组件，其中包含来自 Fabric React 的标签和 PrimaryButton。创建 `ButtonPrimaryExample`：

1. 打开 Yeoman 生成器创建的项目文件夹，并转到 **src\components**。
2. 创建 **button.tsx**。
3. 在 **button.tsx** 中，输入以下代码以创建 `ButtonPrimaryExample` 组件。

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor() {
    super();
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
- 引用用于创建 `ButtonPrimaryExample` 的 Fabric 组件（PrimaryButton、IButtonProps、标签）。
- 使用 `export class ButtonPrimaryExample extends React.Component`，声明并公开新的 `ButtonPrimaryExample` 组件。
- 将 `insertText` 函数声明为处理 `onClick` 事件。
- 在 `render` 函数中定义 React 组件的 UI。呈现器定义组件结构。在 `render` 中，使用 `this.insertText` 连接 `onClick` 事件。

## <a name="step-3---add-the-react-component-to-your-add-in"></a>第 3 步 - 将 React 组件添加到加载项

通过打开 **src\components\app.tsx** 并执行下列操作将 `ButtonPrimaryExample` 添加到外接程序：

- 添加以下导入语句以引用来自步骤 2 中创建的 **button.tsx** 的引用 `ButtonPrimaryExample`（不需要文件扩展名）。

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- 将默认 `render()` 函数替换为以下使用 `<ButtonPrimaryExample />` 的代码。

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

保存所做的更改。所有打开的浏览器实例（包括外接程序）将自动更新和显示 `ButtonPrimaryExample` React 组件。请注意，默认文本和按钮将替换为 `ButtonPrimaryExample` 中定义的文本和主按钮。

## <a name="recommended-components"></a>建议使用的组件

下面列出了建议用于加载项的 Fabric React 用户体验组件：

- [痕迹导航](breadcrumb.md)
- [按钮](button.md)
- [复选框](checkbox.md)
- [ChoiceGroup](choicegroup.md)
- [下拉列表](dropdown.md)
- [标签](label.md)
- [列表](list.md)
- [透视](pivot.md)
- [TextField](textfield.md)
- [切换](toggle.md)

> [!NOTE]
> 今后，我们将陆续添加其他组件。

## <a name="see-also"></a>另请参阅

- [Office UI Fabric React](https://dev.office.com/fabric#/)
- [Fabric React 代码示例入门](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [用户体验设计模式（使用 Fabric 2.6.1）](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office 外接程序 Fabric UI 示例（使用 Fabric 1.0）](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [在 Office 加载项中使用 Fabric 2.6.1](ui-elements/using-office-ui-fabric.md)
- [Office 的 yeoman 生成器](https://github.com/OfficeDev/generator-office)
