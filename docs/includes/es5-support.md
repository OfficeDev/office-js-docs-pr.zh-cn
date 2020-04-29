对于[某些版本的 Office 和 Windows](../concepts/browsers-used-by-office-web-add-ins.md)，运行外接程序的 JavaScript 引擎由 Internet Explorer 提供。 Internet Explorer 引擎不支持高于 ES5 的 JavaScript 版本。 这意味着，如果没有特殊处理，加载项所使用的 JavaScript 文件不能使用在 ES5 后添加到语言中的语法、类型或方法。 这并不意味着必须以 ES5 语法*编写*。 有两个其他选项：

- 在[ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) （也称为 ES6）或更高版本 JavaScript 中或在 TypeScript 中编写代码，然后使用编译器（如[babel](https://babeljs.io/)或[tsc](https://www.typescriptlang.org/index.html)）将代码编译为 ES5 JavaScript。
- 在 ECMAScript 2015 或更高版本的 JavaScript 中编写，但还要加载一个[polyfill](https://wikipedia.org/wiki/Polyfill_(programming))库，如[core-JS](https://github.com/zloirock/core-js) ，使 IE 能够运行您的代码。
