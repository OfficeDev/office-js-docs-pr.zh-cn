---
title: 在 Office 加载项中使用 Office UI Fabric React
description: 了解如何在 Office 加载项中使用 Office UI Fabric React。
ms.date: 09/09/2020
localization_priority: Normal
ms.openlocfilehash: c1671fdd2ea616398f9c57f76898a6cc96daf502
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131981"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="b6cc2-103">在 Office 加载项中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="b6cc2-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="b6cc2-p101">Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果使用 React 生成外接程序，请考虑使用 Fabric React 来创建用户体验。Fabric 提供了多个可在外接程序中使用的基于 React 的 UX 组件，如按钮或复选框。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="b6cc2-107">本文介绍如何创建使用 React 构建的加载项, 并使用 Fabric React 组件。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span>

> [!NOTE]
> <span data-ttu-id="b6cc2-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors)是 Fabric React 附带的，这意味着在完成本文中的步骤后，你的加载项也有权访问 Fabric Core。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="b6cc2-109">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="b6cc2-109">Create an add-in project</span></span>

<span data-ttu-id="b6cc2-110">将使用 Office 加载项的 Yeoman 生成器创建使用 React 的加载项项目。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="b6cc2-111">安装必备组件</span><span class="sxs-lookup"><span data-stu-id="b6cc2-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="b6cc2-112">创建项目</span><span class="sxs-lookup"><span data-stu-id="b6cc2-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="b6cc2-113">**选择项目类型:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="b6cc2-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="b6cc2-114">**选择脚本类型:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="b6cc2-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="b6cc2-115">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="b6cc2-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="b6cc2-116">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="b6cc2-116">**Which Office client application would you like to support?**</span></span> `Word`

![在命令行界面中显示 Yeoman 生成器的提示和答案的屏幕截图](../images/yo-office-word-react.png)

<span data-ttu-id="b6cc2-118">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="b6cc2-119">试用</span><span class="sxs-lookup"><span data-stu-id="b6cc2-119">Try it out</span></span>

1. <span data-ttu-id="b6cc2-120">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="b6cc2-121">完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b6cc2-122">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="b6cc2-123">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="b6cc2-124">你可能还必须以管理员身份运行命令提示符或终端才能进行更改。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="b6cc2-125">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="b6cc2-126">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="b6cc2-127">若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="b6cc2-128">这将启动本地的 Web 服务器（如果尚未运行的话），并使用加载的加载项打开 Word。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="b6cc2-129">若要在浏览器版 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="b6cc2-130">如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="b6cc2-131">若要使用加载项，请在 Word 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="b6cc2-132">在 Word 中，依次选择“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-132">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="b6cc2-133">请注意任务窗格底部的“默认文本”和 "**运行**" 按钮。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="b6cc2-134">在本演练的其余部分中, 你将通过创建使用来自 Fabric React 的 UX 组件的 React 组件来重新定义此文本和按钮。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![显示 Word 应用程序的屏幕截图，其中突出显示了 "显示任务窗格" 功能区按钮，在任务窗格中突出显示了 "运行" 按钮和紧邻文本](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="b6cc2-136">创建使用 Fabric React 的 React 组件</span><span class="sxs-lookup"><span data-stu-id="b6cc2-136">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="b6cc2-137">此时, 你已经创建了一个使用 React 构建的非常基本的任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="b6cc2-138">接下来，完成以下步骤，在加载项项目中创建新的 React 组件 (`ButtonPrimaryExample`)。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="b6cc2-139">该组件使用 Fabric React 的`Label`和`PrimaryButton`组件。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-139">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="b6cc2-140">打开 Yeoman 生成器创建的项目文件夹，并转到 **src\taskpane\components**。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="b6cc2-141">在该文件夹中，创建一个名为“**Button.tsx**”的新文件。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="b6cc2-142">在 **Button.tsx** 中，输入以下代码以定义`ButtonPrimaryExample`组件。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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

<span data-ttu-id="b6cc2-143">此代码将执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="b6cc2-143">This code does the following:</span></span>

- <span data-ttu-id="b6cc2-144">引用使用 `import * as React from 'react';` 的 React 库。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="b6cc2-145">参考用于创建 `ButtonPrimaryExample` 的 Fabric 组件（`PrimaryButton`、`IButtonProps`、`Label`）。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-145">References the Fabric components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="b6cc2-146">声明新的 `ButtonPrimaryExample` 组件使用 `export class ButtonPrimaryExample extends React.Component`。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="b6cc2-147">声明 `insertText` 将处理按钮 `onClick` 事件的函数。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="b6cc2-148">定义 `render` 函数中 React 组件的 UI。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="b6cc2-149">HTML 标记使用 Fabric Reac 中的组件 `Label` 和 `PrimaryButton`，并指定当 `onClick` 事件触发时，`insertText` 函数将运行。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="b6cc2-150">将 React 组件添加到加载项</span><span class="sxs-lookup"><span data-stu-id="b6cc2-150">Add the React component to your add-in</span></span>

<span data-ttu-id="b6cc2-151">通过打开 **src\components\App.tsx** 并完成下列步骤，将组件 `ButtonPrimaryExample` 添加到加载项：</span><span class="sxs-lookup"><span data-stu-id="b6cc2-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="b6cc2-152">添加下面导入语句，以`ButtonPrimaryExample`从 **Button.tsx** 中引用。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="b6cc2-153">删除以下两个导入语句。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="b6cc2-154">将默认 `render()` 函数替换为以下使用 `ButtonPrimaryExample` 的代码。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

4. <span data-ttu-id="b6cc2-155">将所做的更改保存到 **App.tsx**。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="b6cc2-156">查看结果</span><span class="sxs-lookup"><span data-stu-id="b6cc2-156">See the result</span></span>

<span data-ttu-id="b6cc2-157">在 Word 中, 当你保存对 **App.tsx** 的更改时，加载项任务窗格会自动更新。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="b6cc2-158">任务窗格底部的默认文本和按钮现在显示由该`ButtonPrimaryExample`组件定义的 UI。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="b6cc2-159">选择 **插入文本……** 按钮将文本插入到文档中。</span><span class="sxs-lookup"><span data-stu-id="b6cc2-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![显示包含 "插入文本 ..." 的 Word 应用程序的屏幕截图按钮和紧接在文本前面突出显示](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="b6cc2-161">恭喜，您已使用 React 和 Office UI Fabric React 成功创建了任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="b6cc2-161">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span>

## <a name="see-also"></a><span data-ttu-id="b6cc2-162">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b6cc2-162">See also</span></span>

- [<span data-ttu-id="b6cc2-163">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="b6cc2-163">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="b6cc2-164">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="b6cc2-164">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="b6cc2-165">适用于 Office 加载项的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="b6cc2-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="b6cc2-166">Fabric React 代码示例入门</span><span class="sxs-lookup"><span data-stu-id="b6cc2-166">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
