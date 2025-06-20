"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const index_1 = __importStar(require("./index"));
const general_helper_1 = require("./helper/general-helper");
const run = () => __awaiter(void 0, void 0, void 0, function* () {
    const outputDir = `${__dirname}/../__tests__/pptx-output`;
    const templateDir = `${__dirname}/../__tests__/pptx-templates`;
    const automizer = new index_1.default({
        templateDir,
        outputDir
    });
    const pres = automizer
        .loadRoot(`RootTemplate.pptx`)
        .load(`ChartBarsStacked.pptx`, 'charts');
    const dataSmaller = {
        series: [
            { label: 'series s1' },
            { label: 'series s2' }
        ],
        categories: [
            { label: 'item test r1', values: [10, null] },
            { label: 'item test r2', values: [12, 18] },
        ],
    };
    const result = yield pres
        .addSlide('charts', 1, (slide) => {
        slide.modifyElement('BarsStacked', [
            index_1.modify.setChartData(dataSmaller),
        ]);
    })
        .write(`modify-chart-stacked-bars.test.pptx`);
    (0, general_helper_1.vd)(result);
});
run().catch((error) => {
    console.error(error);
});
//# sourceMappingURL=dev.js.map