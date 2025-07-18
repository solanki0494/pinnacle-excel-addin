/*
 * ATTENTION: The "eval" devtool has been used (maybe by default in mode: "development").
 * This devtool is neither made for production nor for readable output files.
 * It uses "eval()" calls to create a separate source file in the browser devtools.
 * If you are trying to read the output file, select a different devtool (https://webpack.js.org/configuration/devtool/)
 * or disable the default devtool with "devtool: false".
 * If you are looking for production-ready output files, see mode: "production" (https://webpack.js.org/configuration/mode/).
 */
/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/commands/commands.ts":
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {

eval("{\n/* global console, Excel, Office */\nObject.defineProperty(exports, \"__esModule\", ({ value: true }));\nconst operatingExpensesService_1 = __webpack_require__(/*! ../services/operatingExpensesService */ \"./src/services/operatingExpensesService.ts\");\n// Initialize Office.js when ready\nfunction initializeOffice() {\n    if (typeof Office !== 'undefined' && Office.onReady) {\n        Office.onReady(() => {\n            console.log('Office.js is ready for commands');\n            // Office is ready - commands will be available\n        });\n    }\n    else {\n        console.log('Office.js not ready for commands, retrying...');\n        setTimeout(initializeOffice, 100);\n    }\n}\n// Start initialization\ninitializeOffice();\nasync function runCalculation(event) {\n    try {\n        await Excel.run(async (context) => {\n            console.log(\"Starting calculation from ribbon button...\");\n            context.application.calculationMode = Excel.CalculationMode.automatic;\n            context.application.calculate(Excel.CalculationType.full);\n            await context.sync();\n            await new Promise(resolve => setTimeout(resolve, 500));\n            await (0, operatingExpensesService_1.copyOperatingExpenses)(context);\n            await (0, operatingExpensesService_1.showNotification)(\"Calculation Complete\", \"Operating expenses copied successfully!\");\n            event.completed();\n        });\n    }\n    catch (error) {\n        console.error(\"Error in runCalculation:\", error);\n        const errorMessage = error instanceof Error ? error.message : String(error);\n        await (0, operatingExpensesService_1.showNotification)(\"Calculation Error\", `Error: ${errorMessage}`);\n        event.completed();\n    }\n}\n__webpack_require__.g.runCalculation = runCalculation;\n\n\n//# sourceURL=webpack://pinnacle-real-estate-excel-addin/./src/commands/commands.ts?\n}");

/***/ }),

/***/ "./src/services/operatingExpensesService.ts":
/*!**************************************************!*\
  !*** ./src/services/operatingExpensesService.ts ***!
  \**************************************************/
/***/ ((__unused_webpack_module, exports) => {

eval("{\nObject.defineProperty(exports, \"__esModule\", ({ value: true }));\nexports.copyOperatingExpenses = copyOperatingExpenses;\nexports.showNotification = showNotification;\nasync function copyOperatingExpenses(context) {\n    try {\n        console.log(\"Starting operating expenses copy...\");\n        const outputsSheet = context.workbook.worksheets.getItem(\"Outputs\");\n        const engineeringSheet = context.workbook.worksheets.getItem(\"Software Engineer Cash Flow\");\n        const sourceRange = outputsSheet.getRange(\"O32:CI35\");\n        const targetRange = engineeringSheet.getRange(\"O32:CI35\");\n        sourceRange.load(\"values\");\n        await context.sync();\n        targetRange.values = sourceRange.values;\n        await context.sync();\n        console.log(\"Operating expenses copied successfully from Outputs to Engineering sheet\");\n    }\n    catch (error) {\n        console.error(\"Error copying operating expenses:\", error);\n        throw error;\n    }\n}\nasync function showNotification(title, message) {\n    try {\n        await Excel.run(async (context) => {\n            const notification = {\n                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,\n                message: message,\n                icon: \"icon1\",\n                persistent: false\n            };\n            console.log(`${title}: ${message}`);\n        });\n    }\n    catch (error) {\n        console.log(`${title}: ${message}`);\n    }\n}\n\n\n//# sourceURL=webpack://pinnacle-real-estate-excel-addin/./src/services/operatingExpensesService.ts?\n}");

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	})();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module can't be inlined because the eval devtool is used.
/******/ 	var __webpack_exports__ = __webpack_require__("./src/commands/commands.ts");
/******/ 	
/******/ })()
;