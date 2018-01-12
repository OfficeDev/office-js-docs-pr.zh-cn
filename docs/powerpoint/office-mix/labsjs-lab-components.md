
# <a name="labsjs-lab-components"></a>LabsJS lab components

Lab.js 为你提供了四种可用来装配你的实验室的组件类型。每个组件类型支持某个特定类型的实验室交互，包括：课程 HTML 格式的 iFrame 中的多个选择问题、免费响应问题或类似查看 Web 页面的活动。

## <a name="components"></a>组件

Office Mix 支持以下四种实验室组件类型： 


-  **Activity component** ( **IActivityComponent**). Presents the user with an activity that must be completed; for example, read a piece of text, watch a video, or interact with a simulation. For more information, see [Labs.Components.ActivityComponentInstance](../../../reference/office-mix/labs.components.activitycomponentinstance.md).
    
-  **Choice component** ( **IChoiceComponent**). Presents the user with a list of choices from which the user must select. Supports single or multiple responses (or no answer at all). Use this component type for true/false, multiple choice, multiple response, or polls. For more information, see [Labs.Components.ChoiceComponentInstance](../../../reference/office-mix/labs.components.choicecomponentinstance.md).
    
-  **Input component** ( **IInputComponent**). Enables free form user input. Use this component type when you want to get responses to questions or math problems from the user, for example, or for other problem types that require text inputs from the user. For more information, see [Labs.Components.InputComponentInstance](../../../reference/office-mix/labs.components.inputcomponentinstance.md).
    
-  **Dynamic component** ( **IDynamicComponent**). Generates other component types at runtime. Use this component type when you have branching questions, for example, where follow-up component types vary depending on a previous user input. This type also enables creating quiz banks or generating problems at runtime. For more information, see [Labs.Components.DynamicComponentInstance](../../../reference/office-mix/labs.components.dynamiccomponentinstance.md).
    

## <a name="additional-resources"></a>其他资源



- [Office Mix 外接程序](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [配置和编辑适用于 Office Mix 的 LabsJS 实验室](../../powerpoint/office-mix/configuring-and-editing-labsjs-labs-for-office-mix.md)
    
- [演练︰创建第一个 Office Mix 实验室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
