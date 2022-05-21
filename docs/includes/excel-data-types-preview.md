> [!NOTE]
> 数据类型 API 目前仅在公共预览版中提供。 预览 API 可能会发生变更，不适合在生产环境中使用。 我们建议你仅在测试和开发环境中试用它们。 不要在生产环境或业务关键型文档中使用预览 API。
>
> 若要使用预览 API：
>
> - 必须在内容分发网络 （CDN） （https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)） 上引用 **beta** 库。 用于 TypeScript 编译和 IntelliSense 的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)位于 CDN 和 [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) 中。 可以使用 `npm install --save-dev @types/office-js-preview` 来安装这些类型。 有关其他信息，请参阅 [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM 包自述文件。
> - 可能需要加入 [Office 预览体验计划](https://insider.office.com)才能访问更新的 Office 版本。
>
> 若要在 Windows 版 Office 中试用数据类型，则 Excel 内部版本号必须大于或等于 16.0.14626.10000。 若要尝试 Mac 版 Office 中的数据类型集成，Excel 内部版本号必须大于或等于 16.55.21102600。