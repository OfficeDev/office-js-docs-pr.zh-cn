---
title: '? Office ?????? Office UI Fabric React'
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8ae8bac8c8043b51188d765dd7170922dcc1c84e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="98bf1-102">? Office ?????? Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="98bf1-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="98bf1-p101">Office UI Fabric ????? Office ? Office 365 ????? JavaScript ????????? React ???????????? Fabric React ????????Fabric ????????????????? React ? UX ???????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="98bf1-106">??????????? Fabric React ???????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="98bf1-107">??????????????????????? Fabric Core?</span><span class="sxs-lookup"><span data-stu-id="98bf1-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="98bf1-108">? 1 ? - ????? Office ? Yeoman ???????</span><span class="sxs-lookup"><span data-stu-id="98bf1-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="98bf1-p102">?????? Fabric React ???????????? Office ? Yeoman ????Office ? Yeoman ??????? Office ?????????????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-p102">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office. The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office add-in.</span></span>

<span data-ttu-id="98bf1-111">?????????? **Windows PowerShell**?????????????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="98bf1-112">???????</span><span class="sxs-lookup"><span data-stu-id="98bf1-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="98bf1-113">?? `yo office`?????????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="98bf1-114">?????????? Office ???????????? **Word**?</span><span class="sxs-lookup"><span data-stu-id="98bf1-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="98bf1-p103">?????????????????? `npm start`?????????????????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="98bf1-117">[?????](..\testing\test-debug-office-add-ins.md)?????????? UI?</span><span class="sxs-lookup"><span data-stu-id="98bf1-117">[Sideload your manifest](..\testing\test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="98bf1-118">? 2 ? - ?? Fabric React ??</span><span class="sxs-lookup"><span data-stu-id="98bf1-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="98bf1-p104">????? Fabric React ?????????????? `ButtonPrimaryExample` ??? React ????????? Fabric React ???? PrimaryButton??? `ButtonPrimaryExample`?</span><span class="sxs-lookup"><span data-stu-id="98bf1-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="98bf1-122">?? Yeoman ??????????????? **src\components**?</span><span class="sxs-lookup"><span data-stu-id="98bf1-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="98bf1-123">?? **button.tsx**?</span><span class="sxs-lookup"><span data-stu-id="98bf1-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="98bf1-124">? **button.tsx** ??????????? `ButtonPrimaryExample` ???</span><span class="sxs-lookup"><span data-stu-id="98bf1-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

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

<span data-ttu-id="98bf1-125">???????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-125">This code does the following:</span></span>

- <span data-ttu-id="98bf1-126">???? `import * as React from 'react';` ? React ??</span><span class="sxs-lookup"><span data-stu-id="98bf1-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="98bf1-127">?????? `ButtonPrimaryExample` ? Fabric ???PrimaryButton?IButtonProps?????</span><span class="sxs-lookup"><span data-stu-id="98bf1-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="98bf1-128">?? `export class ButtonPrimaryExample extends React.Component`???????? `ButtonPrimaryExample` ???</span><span class="sxs-lookup"><span data-stu-id="98bf1-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="98bf1-129">? `insertText` ??????? `onClick` ???</span><span class="sxs-lookup"><span data-stu-id="98bf1-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="98bf1-p105">? `render` ????? React ??? UI???????????? `render` ???? `this.insertText` ?? `onClick` ???</span><span class="sxs-lookup"><span data-stu-id="98bf1-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="98bf1-133">? 3 ? - ? React ????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="98bf1-134">???? **src\components\app.tsx** ???????? `ButtonPrimaryExample` ????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="98bf1-135">??????????????? 2 ???? **button.tsx** ??? `ButtonPrimaryExample`???????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="98bf1-136">??? `render()` ????????? `<ButtonPrimaryExample />` ????</span><span class="sxs-lookup"><span data-stu-id="98bf1-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

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

<span data-ttu-id="98bf1-p106">?????????????????????????????????? `ButtonPrimaryExample` React ?????????????????? `ButtonPrimaryExample` ???????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>

## <a name="recommended-components"></a><span data-ttu-id="98bf1-140">???????</span><span class="sxs-lookup"><span data-stu-id="98bf1-140">Recommended components</span></span>

<span data-ttu-id="98bf1-141">????????????? Fabric React ???????</span><span class="sxs-lookup"><span data-stu-id="98bf1-141">The following is a list of the Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="98bf1-142">????</span><span class="sxs-lookup"><span data-stu-id="98bf1-142">Breadcrumb</span></span>](breadcrumb.md)
- [<span data-ttu-id="98bf1-143">??</span><span class="sxs-lookup"><span data-stu-id="98bf1-143">Button</span></span>](button.md)
- [<span data-ttu-id="98bf1-144">???</span><span class="sxs-lookup"><span data-stu-id="98bf1-144">Checkbox</span></span>](checkbox.md)
- [<span data-ttu-id="98bf1-145">???</span><span class="sxs-lookup"><span data-stu-id="98bf1-145">ChoiceGroup</span></span>](choicegroup.md)
- [<span data-ttu-id="98bf1-146">????</span><span class="sxs-lookup"><span data-stu-id="98bf1-146">Dropdown</span></span>](dropdown.md)
- [<span data-ttu-id="98bf1-147">??</span><span class="sxs-lookup"><span data-stu-id="98bf1-147">Label</span></span>](label.md)
- [<span data-ttu-id="98bf1-148">??</span><span class="sxs-lookup"><span data-stu-id="98bf1-148">List</span></span>](list.md)
- [<span data-ttu-id="98bf1-149">??</span><span class="sxs-lookup"><span data-stu-id="98bf1-149">Pivot</span></span>](pivot.md)
- [<span data-ttu-id="98bf1-150">???</span><span class="sxs-lookup"><span data-stu-id="98bf1-150">TextField</span></span>](textfield.md)
- [<span data-ttu-id="98bf1-151">??</span><span class="sxs-lookup"><span data-stu-id="98bf1-151">Toggle</span></span>](toggle.md)

> [!NOTE]
> <span data-ttu-id="98bf1-152">???????????????</span><span class="sxs-lookup"><span data-stu-id="98bf1-152">We will add additional components over time.</span></span>

## <a name="see-also"></a><span data-ttu-id="98bf1-153">????</span><span class="sxs-lookup"><span data-stu-id="98bf1-153">See also</span></span>

- [<span data-ttu-id="98bf1-154">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="98bf1-154">Office UI Fabric React</span></span>](https://dev.office.com/fabric#/)
- [<span data-ttu-id="98bf1-155">Fabric React ??????</span><span class="sxs-lookup"><span data-stu-id="98bf1-155">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="98bf1-156">??????????? Fabric 2.6.1?</span><span class="sxs-lookup"><span data-stu-id="98bf1-156">UX design patterns (uses Fabric 2.6.1)</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="98bf1-157">Office ???? Fabric UI ????? Fabric 1.0?</span><span class="sxs-lookup"><span data-stu-id="98bf1-157">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="98bf1-158">Office ? Yeoman ???</span><span class="sxs-lookup"><span data-stu-id="98bf1-158">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
