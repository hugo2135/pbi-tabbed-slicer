/**
 * settings.ts — 格式窗格設定模型（powerbi-visuals-utils-formattingmodel）
 */

import powerbi from "powerbi-visuals-api";
import ValidatorType = powerbi.visuals.ValidatorType;

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { MAX_TABS } from "./dataModel";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;

/**
 * 預設頁籤配色 —— 依頁籤總數把色相平均分配，不用維護寫死的色票。
 * 這樣 `npm run set-tabs` 改幾個頁籤，就自動有幾個不重複的顏色。
 *
 * 飽和度與亮度刻意壓低，讓底色柔和、文字好讀，也不會跟強調色打架。
 */
const TAB_COLOR_SATURATION = 0.45;
const TAB_COLOR_LIGHTNESS = 0.88;
/** 起始色相：從藍綠開始，避開一開頭就是大紅 */
const TAB_COLOR_HUE_OFFSET = 195;

function hslToHex(h: number, s: number, l: number): string {
    const hue = ((h % 360) + 360) % 360;
    const c = (1 - Math.abs(2 * l - 1)) * s;
    const x = c * (1 - Math.abs(((hue / 60) % 2) - 1));
    const m = l - c / 2;
    let rgb: [number, number, number];
    if (hue < 60) { rgb = [c, x, 0]; }
    else if (hue < 120) { rgb = [x, c, 0]; }
    else if (hue < 180) { rgb = [0, c, x]; }
    else if (hue < 240) { rgb = [0, x, c]; }
    else if (hue < 300) { rgb = [x, 0, c]; }
    else { rgb = [c, 0, x]; }
    return "#" + rgb
        .map(v => Math.round((v + m) * 255).toString(16).padStart(2, "0").toUpperCase())
        .join("");
}

/** 第 index 個頁籤（0-based）的預設背景色 */
export function defaultTabColor(index: number, total: number = MAX_TABS): string {
    const count = Math.max(1, total);
    const hue = TAB_COLOR_HUE_OFFSET + (360 * index) / count;
    return hslToHex(hue, TAB_COLOR_SATURATION, TAB_COLOR_LIGHTNESS);
}

class TabsCard extends Card {
    name: string = "tabs";
    displayName: string = "頁籤";

    perPage = new formattingSettings.NumUpDown({
        name: "perPage",
        displayName: "每頁顯示頁籤數上限",
        value: 5,
        options: {
            minValue: { type: ValidatorType.Min, value: 2 },
            maxValue: { type: ValidatorType.Max, value: MAX_TABS }
        }
    });

    autoName = new formattingSettings.ToggleSwitch({
        name: "autoName",
        displayName: "自動使用欄位名稱",
        value: true
    });

    useTabColors = new formattingSettings.ToggleSwitch({
        name: "useTabColors",
        displayName: "啟用頁籤色彩",
        value: false
    });

    names: formattingSettings.TextInput[] = [];
    colors: formattingSettings.ColorPicker[] = [];

    slices = [];

    constructor() {
        super();
        for (let i = 1; i <= MAX_TABS; i++) {
            this.names.push(new formattingSettings.TextInput({
                name: `name${i}`,
                displayName: `頁籤 ${i} 名稱`,
                value: "",
                placeholder: "留空則使用欄位名稱"
            }));
            this.colors.push(new formattingSettings.ColorPicker({
                name: `bg${i}`,
                displayName: `頁籤 ${i} 背景色`,
                value: { value: defaultTabColor(i - 1) }
            }));
        }
        this.slices = [this.perPage, this.autoName, this.useTabColors, ...this.names, ...this.colors];
    }

    /** 只有啟用頁籤色彩時才顯示各頁籤的色彩選擇器 */
    public syncVisibility(): void {
        this.colors.forEach(c => { c.visible = this.useTabColors.value; });
    }
}


class BehaviorCard extends Card {
    name: string = "behavior";
    displayName: string = "篩選行為";

    mode = new formattingSettings.ItemDropdown({
        name: "mode",
        displayName: "篩選模式",
        items: [
            { value: "filter", displayName: "篩選器 (Filter)" },
            { value: "select", displayName: "選取 / 交叉醒目提示 (Select)" }
        ],
        value: { value: "filter", displayName: "篩選器 (Filter)" }
    });

    singleSelect = new formattingSettings.ToggleSwitch({
        name: "singleSelect",
        displayName: "單選模式",
        value: false
    });

    selectChildren = new formattingSettings.ToggleSwitch({
        name: "selectChildren",
        displayName: "勾選父項時連動子項",
        value: true
    });

    showSelectAll = new formattingSettings.ToggleSwitch({
        name: "showSelectAll",
        displayName: "顯示「全選」列",
        value: true
    });

    expandAll = new formattingSettings.ToggleSwitch({
        name: "expandAll",
        displayName: "預設全部展開",
        value: false
    });

    resetOnTabChange = new formattingSettings.ToggleSwitch({
        name: "resetOnTabChange",
        displayName: "切換頁籤時清除搜尋字串",
        value: true
    });

    crossTab = new formattingSettings.ToggleSwitch({
        name: "crossTab",
        displayName: "頁籤之間互相篩選",
        value: true
    });

    slices = [
        this.mode,
        this.crossTab,
        this.singleSelect,
        this.selectChildren,
        this.showSelectAll,
        this.expandAll,
        this.resetOnTabChange
    ];
}

class HeaderCard extends Card {
    name: string = "header";
    displayName: string = "標題";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "顯示",
        value: true
    });

    topLevelSlice = this.show;

    text = new formattingSettings.TextInput({
        name: "text",
        displayName: "標題文字",
        value: "篩選條件",
        placeholder: "篩選條件"
    });

    showCount = new formattingSettings.ToggleSwitch({
        name: "showCount",
        displayName: "顯示已選取數量",
        value: true
    });

    showHelp = new formattingSettings.ToggleSwitch({
        name: "showHelp",
        displayName: "顯示說明按鈕",
        value: true
    });

    helpUrl = new formattingSettings.TextInput({
        name: "helpUrl",
        displayName: "說明文件網址 (選填)",
        value: "",
        placeholder: "https://..."
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "字型色彩",
        value: { value: "#201E1D" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "字型大小",
        value: 14
    });

    slices = [this.text, this.showCount, this.showHelp, this.helpUrl, this.fontColor, this.fontSize];
}

class SearchCard extends Card {
    name: string = "search";
    displayName: string = "搜尋";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "顯示搜尋框",
        value: true
    });

    topLevelSlice = this.show;

    serverSide = new formattingSettings.ToggleSwitch({
        name: "serverSide",
        displayName: "伺服器端搜尋",
        value: false
    });

    matchLevel = new formattingSettings.NumUpDown({
        name: "matchLevel",
        displayName: "比對層級（0 = 最底層）",
        value: 0,
        options: {
            minValue: { type: ValidatorType.Min, value: 0 },
            maxValue: { type: ValidatorType.Max, value: 10 }
        }
    });

    placeholder = new formattingSettings.TextInput({
        name: "placeholder",
        displayName: "提示文字",
        value: "搜尋…",
        placeholder: "搜尋…"
    });

    searchAllLevels = new formattingSettings.ToggleSwitch({
        name: "searchAllLevels",
        displayName: "搜尋所有層級",
        value: true
    });

    slices = [this.placeholder, this.searchAllLevels, this.serverSide, this.matchLevel];
}

class ItemsCard extends Card {
    name: string = "items";
    displayName: string = "項目";

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "字型色彩",
        value: { value: "#201E1D" }
    });

    accentColor = new formattingSettings.ColorPicker({
        name: "accentColor",
        displayName: "強調色彩",
        value: { value: "#0088B0" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "背景色彩",
        value: { value: "#FFFFFF" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "字型大小",
        value: 12
    });

    rowHeight = new formattingSettings.NumUpDown({
        name: "rowHeight",
        displayName: "列高 (px)",
        value: 28
    });

    indent = new formattingSettings.NumUpDown({
        name: "indent",
        displayName: "階層縮排 (px)",
        value: 18
    });

    showCounts = new formattingSettings.ToggleSwitch({
        name: "showCounts",
        displayName: "顯示子項數量",
        value: true
    });

    slices = [
        this.fontColor,
        this.accentColor,
        this.backgroundColor,
        this.fontSize,
        this.rowHeight,
        this.indent,
        this.showCounts
    ];
}

export class VisualSettings extends Model {
    tabs = new TabsCard();
    behavior = new BehaviorCard();
    header = new HeaderCard();
    search = new SearchCard();
    items = new ItemsCard();

    cards = [this.tabs, this.behavior, this.header, this.search, this.items];

    /** 取得使用者自訂的頁籤名稱陣列 */
    get tabNames(): string[] {
        return this.tabs.names.map(n => n.value || "");
    }
}
