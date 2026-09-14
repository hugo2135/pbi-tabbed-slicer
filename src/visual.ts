/**
 * visual.ts — 分頁籤階層篩選器 (Tabbed Hierarchy Slicer)
 * ------------------------------------------------------------------
 * 版面：標題 / 頁籤列（可分頁、可滾輪捲動）/ 搜尋框 / 階層勾選清單
 * 完全響應式：所有尺寸依 update() 拿到的 viewport 重新計算。
 */

import "./../style/visual.less";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { IFilter } from "powerbi-models";

import { VisualSettings } from "./settings";
import { transform, MAX_TABS, SlicerModel, TabModel, TreeNode } from "./dataModel";
import { FilterEngine, SelectionStore, TabSelection, collectSelection, pathSignature } from "./filterEngine";

/** #rrggbb / #rgb → [r,g,b]；解析失敗回傳 null */
function parseHex(hex: string): [number, number, number] | null {
    const s = (hex || "").trim().replace("#", "");
    if (s.length === 3) {
        const r = parseInt(s[0] + s[0], 16);
        const g = parseInt(s[1] + s[1], 16);
        const b = parseInt(s[2] + s[2], 16);
        return isNaN(r + g + b) ? null : [r, g, b];
    }
    if (s.length === 6) {
        const r = parseInt(s.substring(0, 2), 16);
        const g = parseInt(s.substring(2, 4), 16);
        const b = parseInt(s.substring(4, 6), 16);
        return isNaN(r + g + b) ? null : [r, g, b];
    }
    return null;
}

/** amount > 0 往白色靠，< 0 往黑色靠 */
function shade(hex: string, amount: number): string {
    const rgb = parseHex(hex);
    if (!rgb) {
        return hex;
    }
    const mix = (v: number): number => amount >= 0
        ? Math.round(v + (255 - v) * amount)
        : Math.round(v * (1 + amount));
    const toHex = (v: number): string => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0");
    return "#" + rgb.map(mix).map(toHex).join("");
}

/** 依背景亮度挑深色或淺色文字，確保可讀 */
function contrastText(hex: string): string {
    const rgb = parseHex(hex);
    if (!rgb) {
        return "";
    }
    // 相對亮度（sRGB 近似）
    const [r, g, b] = rgb;
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.6 ? "#201E1D" : "#FFFFFF";
}

/** 清空節點內容（避免使用 innerHTML，符合自訂視覺的安全性檢查） */
function clearElement(el: HTMLElement): void {
    while (el.firstChild) {
        el.removeChild(el.firstChild);
    }
}

/** 說明面板的內容。改功能時記得一起更新。 */
const HELP_SECTIONS: { title: string; items: string[] }[] = [
    {
        title: "選取",
        items: [
            "點項目或勾選框即可篩選，可同時勾多個。",
            "按住 Shift 點第一項設為起點（左側出現色條），再按住 Shift 點第二項，中間全部一起勾選。",
            "一般點擊不會影響 Shift 的起點。",
            "勾選父項會連同底下所有子項一起選取。"
        ]
    },
    {
        title: "展開 / 折疊",
        items: [
            "項目左側的三角形可展開下一層。",
            "在有下層的項目上按滑鼠右鍵，也可以直接展開或折疊。",
            "上方「全部展開 / 全部折疊」只作用在目前的頁籤。"
        ]
    },
    {
        title: "搜尋",
        items: [
            "輸入關鍵字會同時比對所有層級，命中的分支會自動展開。",
            "搜尋時「全選」會變成「全選搜尋結果」，只勾選符合的項目。",
            "資料量大時可請報表製作者開啟「伺服器端搜尋」，就能搜到超過顯示上限之外的項目（搜尋框閃動代表正在查詢）。"
        ]
    },
    {
        title: "頁籤",
        items: [
            "每個頁籤是一組獨立的篩選條件，彼此以 AND 結合。",
            "頁籤放不下時會出現 ‹ › 按鈕，也可以用滑鼠滾輪橫向捲動。",
            "在某個頁籤勾選後，其他頁籤的清單會跟著縮減。"
        ]
    },
    {
        title: "清除",
        items: [
            "點右上「已選 N 項」會清除**目前頁籤**的勾選。",
            "其他頁籤也有勾選時，右邊會出現「全部 N · 清除」，點它才是清掉所有頁籤。"
        ]
    }
];

interface Row {
    node: TreeNode;
    depth: number;
    hasChildren: boolean;
    expanded: boolean;
    checked: boolean;
    indeterminate: boolean;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private formattingService: FormattingSettingsService;
    private settings: VisualSettings;
    private filterEngine: FilterEngine;

    private root: HTMLElement;
    private headerEl: HTMLElement;
    private titleEl: HTMLElement;
    private countEl: HTMLElement;
    private helpBtn: HTMLButtonElement;
    private helpPanelEl: HTMLElement;
    private helpOpen = false;
    private tabsEl: HTMLElement;
    private tabsTrackEl: HTMLElement;
    private prevBtn: HTMLButtonElement;
    private nextBtn: HTMLButtonElement;
    private toolbarEl: HTMLElement;
    private expandAllBtn: HTMLButtonElement;
    private collapseAllBtn: HTMLButtonElement;
    private searchWrapEl: HTMLElement;
    private searchEl: HTMLInputElement;
    private statusBarEl: HTMLElement;
    private selectAllBox: HTMLInputElement;
    private selectAllLabel: HTMLElement;
    private listEl: HTMLElement;
    private emptyEl: HTMLElement;

    private model: SlicerModel = { tabs: [], isEmpty: true, rowCount: 0, truncated: false };
    private activeTab = 0;
    private tabStart = 0;
    private query = "";
    private expanded = new Set<string>();
    /** 畫面用：被勾選的節點 key */
    private checked = new Set<string>();
    /**
     * 篩選用：各頁籤已勾選的葉層路徑。
     * 跨頁籤篩選開啟後，其他頁籤的項目會從 dataView 消失，
     * 所以條件必須從這裡組，不能每次從樹重算。
     */
    private selection: SelectionStore = new Map();
    private wheelAcc = 0;
    private viewport = { width: 0, height: 0 };
    /**
     * 按過「全部折疊」後，要能蓋過格式窗格的「預設全部展開」設定，
     * 否則按了沒反應。null = 跟隨設定。
     */
    private expandAllOverride: boolean | null = null;
    /** 目前畫面上顯示的列，Shift 區間勾選要靠它算範圍 */
    private visibleRows: Row[] = [];
    /**
     * Shift 區間的起點。**只有 Shift 點擊會設定它**，一般點擊不會動到 ——
     * 否則使用者先隨手勾了一項，之後想用 Shift 圈範圍時，
     * 起點會莫名其妙變成那個隨手勾的項目。
     * 用 key 而不是索引，這樣重繪、捲動、展開折疊都不會錯位。
     */
    private shiftAnchorKey: string | null = null;
    /** click 先於 change 觸發，用它把 Shift 狀態傳給 change 處理 */
    private shiftPending = false;
    /** 伺服器端搜尋的防抖計時器 */
    private searchTimer: number | null = null;
    /** 已經送進查詢的搜尋字串，用來判斷需不需要重新查 */
    private appliedQuery = "";
    /** 等待查詢回來時顯示提示，避免使用者以為當掉 */
    private searchPending = false;
    /**
     * 還有幾次「我們自己送出的」update 還沒收到。
     *
     * 一次 commit() 可能同時改 general.filter 和 general.selfFilter
     * 兩個獨立屬性（伺服器端搜尋開著時尤其如此），每個屬性各自可能觸發
     * 一次 update()。曾經用布林值只擋得住第一次，第二次就被誤判成外部
     * 帶進來的條件，從 jsonFilters 重建出不完整的勾選狀態，
     * 導致「取消勾選 B 時連 A 也被清掉」之類的假象。所以改成計數器，
     * 送出幾次 applyJsonFilter 就記幾次，每收到一次 update 才扣一次。
     */
    private pendingSelfUpdates = 0;
    private events: powerbi.extensibility.IVisualEventService;
    private isHighContrast = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingService = new FormattingSettingsService();
        this.filterEngine = new FilterEngine(this.host, this.selectionManager);
        this.events = this.host.eventService;
        this.buildDom(options.element);

        // 右鍵內容功能表。
        // 在有下層的項目上按右鍵是「展開/折疊」（由該列自己處理並吃掉事件），
        // 其他地方才叫出 Power BI 的內容功能表。
        this.root.addEventListener("contextmenu", (ev: MouseEvent) => {
            ev.preventDefault();
            if ((ev as MouseEvent & { ttsHandled?: boolean }).ttsHandled) {
                return;
            }
            this.selectionManager.showContextMenu({}, { x: ev.clientX, y: ev.clientY });
        });
    }

    // ── DOM 骨架（只建一次，之後只更新內容） ────────────────────────

    private buildDom(container: HTMLElement): void {
        this.root = document.createElement("div");
        this.root.className = "tts-root";
        container.appendChild(this.root);

        // 標題列
        this.headerEl = document.createElement("div");
        this.headerEl.className = "tts-header";
        this.titleEl = document.createElement("div");
        this.titleEl.className = "tts-title";
        this.countEl = document.createElement("div");
        this.countEl.className = "tts-count";
        this.countEl.setAttribute("role", "status");

        this.helpBtn = document.createElement("button");
        this.helpBtn.type = "button";
        this.helpBtn.className = "tts-help";
        this.helpBtn.textContent = "?";
        this.helpBtn.title = "操作說明";
        this.helpBtn.setAttribute("aria-label", "操作說明");
        this.helpBtn.addEventListener("click", (ev) => {
            ev.stopPropagation();
            this.toggleHelp();
        });

        this.headerEl.appendChild(this.titleEl);
        this.headerEl.appendChild(this.countEl);
        this.headerEl.appendChild(this.helpBtn);
        this.root.appendChild(this.headerEl);

        // 說明面板（覆蓋在清單上，預設隱藏）
        this.helpPanelEl = document.createElement("div");
        this.helpPanelEl.className = "tts-helppanel";
        this.helpPanelEl.style.display = "none";
        this.root.appendChild(this.helpPanelEl);

        // 點面板以外的地方就關閉
        this.root.addEventListener("click", () => {
            if (this.helpOpen) {
                this.toggleHelp(false);
            }
        });
        this.helpPanelEl.addEventListener("click", ev => ev.stopPropagation());

        // 頁籤列
        this.tabsEl = document.createElement("div");
        this.tabsEl.className = "tts-tabs";
        this.tabsEl.setAttribute("role", "tablist");

        this.prevBtn = document.createElement("button");
        this.prevBtn.type = "button";
        this.prevBtn.className = "tts-tabnav";
        this.prevBtn.setAttribute("aria-label", "上一組頁籤");
        this.prevBtn.textContent = "‹";
        this.prevBtn.addEventListener("click", () => this.scrollTabs(-1));

        this.nextBtn = document.createElement("button");
        this.nextBtn.type = "button";
        this.nextBtn.className = "tts-tabnav";
        this.nextBtn.setAttribute("aria-label", "下一組頁籤");
        this.nextBtn.textContent = "›";
        this.nextBtn.addEventListener("click", () => this.scrollTabs(1));

        this.tabsTrackEl = document.createElement("div");
        this.tabsTrackEl.className = "tts-tabs-track";

        this.tabsEl.appendChild(this.prevBtn);
        this.tabsEl.appendChild(this.tabsTrackEl);
        this.tabsEl.appendChild(this.nextBtn);
        this.root.appendChild(this.tabsEl);

        this.tabsEl.addEventListener("wheel", (ev: WheelEvent) => {
            if (this.model.tabs.length <= this.tabsPerPage()) {
                return;
            }
            ev.preventDefault();
            const delta = Math.abs(ev.deltaY) > Math.abs(ev.deltaX) ? ev.deltaY : ev.deltaX;
            this.wheelAcc += delta;
            if (Math.abs(this.wheelAcc) < 30) {
                return;
            }
            const step = this.wheelAcc > 0 ? 1 : -1;
            this.wheelAcc = 0;
            this.scrollTabs(step);
        }, { passive: false });

        // 工具列（標題／頁籤 與 搜尋列 之間）：全部展開 / 全部折疊
        this.toolbarEl = document.createElement("div");
        this.toolbarEl.className = "tts-toolbar";

        this.expandAllBtn = document.createElement("button");
        this.expandAllBtn.type = "button";
        this.expandAllBtn.className = "tts-toolbtn";
        this.expandAllBtn.textContent = "全部展開";
        this.expandAllBtn.addEventListener("click", () => this.setAllExpanded(true));

        this.collapseAllBtn = document.createElement("button");
        this.collapseAllBtn.type = "button";
        this.collapseAllBtn.className = "tts-toolbtn";
        this.collapseAllBtn.textContent = "全部折疊";
        this.collapseAllBtn.addEventListener("click", () => this.setAllExpanded(false));

        this.toolbarEl.appendChild(this.expandAllBtn);
        this.toolbarEl.appendChild(this.collapseAllBtn);
        this.root.appendChild(this.toolbarEl);

        // 搜尋列
        this.searchWrapEl = document.createElement("div");
        this.searchWrapEl.className = "tts-searchwrap";
        this.searchEl = document.createElement("input");
        this.searchEl.type = "search";
        this.searchEl.className = "tts-search";
        this.searchEl.addEventListener("input", () => {
            this.query = this.searchEl.value;
            // 清單內容變了，原本的區間起點不再有意義
            this.shiftAnchorKey = null;
            this.renderList();
            this.scheduleServerSearch();
        });
        this.searchWrapEl.appendChild(this.searchEl);
        this.root.appendChild(this.searchWrapEl);

        // 狀態列：固定在捲動區之外，永遠存在。
        // 左邊是「全選」（可由格式窗格關閉），右邊是「未篩選／已選 N 項」。
        // 固定住的好處是捲到清單深處也還能全選與清除。
        this.statusBarEl = document.createElement("div");
        this.statusBarEl.className = "tts-statusbar";

        this.selectAllBox = document.createElement("input");
        this.selectAllBox.type = "checkbox";
        this.selectAllBox.className = "tts-check";

        this.selectAllLabel = document.createElement("div");
        this.selectAllLabel.className = "tts-label tts-label-all";
        this.selectAllLabel.addEventListener("click", () => {
            this.selectAllBox.checked = !this.selectAllBox.checked;
            this.selectAllBox.dispatchEvent(new Event("change"));
        });

        this.statusBarEl.appendChild(this.selectAllBox);
        this.statusBarEl.appendChild(this.selectAllLabel);
        this.statusBarEl.appendChild(this.countEl);
        this.root.appendChild(this.statusBarEl);

        // 清單
        this.listEl = document.createElement("div");
        this.listEl.className = "tts-list";
        this.listEl.setAttribute("role", "tree");
        this.root.appendChild(this.listEl);

        this.emptyEl = document.createElement("div");
        this.emptyEl.className = "tts-empty";
        this.root.appendChild(this.emptyEl);
    }

    // ── Power BI 生命週期 ───────────────────────────────────────────

    public update(options: VisualUpdateOptions): void {
        this.events?.renderingStarted(options);
        try {
            this.updateCore(options);
            this.events?.renderingFinished(options);
        } catch (e) {
            this.events?.renderingFailed(options, String(e));
        }
    }

    private updateCore(options: VisualUpdateOptions): void {
        const dataViews = (options.dataViews || []).filter(dv => !!dv);
        // 頁籤 1 可能沒有繫結欄位，因此取第一個實際存在的 dataView 來讀格式設定
        this.settings = this.formattingService.populateFormattingSettingsModel(
            VisualSettings,
            dataViews[0]
        );
        this.viewport = {
            width: options.viewport.width,
            height: options.viewport.height
        };

        const previousTabCount = this.model.tabs.length;
        this.model = transform(
            this.host,
            dataViews[0],
            this.settings.tabNames,
            this.settings.tabs.autoName.value
        );

        // 各頁籤的欄位在同一次查詢中會交叉組合，列數頂到上限時清單可能不完整
        if (this.model.truncated) {
            this.host.displayWarningIcon(
                "資料量已達上限",
                "各頁籤的欄位會在同一次查詢中交叉組合，列數已達 30,000 上限，部分項目可能未列出。請改放低基數的維度欄位。"
            );
        }

        // 只清掉「頁籤本身消失了」的狀態。
        // 不能因為某個值這次沒回傳就把勾選清掉 —— 跨頁籤篩選開啟時，
        // 其他頁籤的項目本來就會被濾掉，清掉會造成條件反覆消失又出現。
        const liveTabs = new Set(this.model.tabs.map(t => t.index));
        const tabOfKey = (key: string): number => {
            const m = /^t(\d+)\|/.exec(key);
            return m ? parseInt(m[1], 10) : -1;
        };
        this.checked.forEach(k => { if (!liveTabs.has(tabOfKey(k))) { this.checked.delete(k); } });
        this.expanded.forEach(k => { if (!liveTabs.has(tabOfKey(k))) { this.expanded.delete(k); } });
        this.selection.forEach((_v, index) => { if (!liveTabs.has(index)) { this.selection.delete(index); } });

        if (this.activeTab >= this.model.tabs.length) {
            this.activeTab = 0;
        }
        if (previousTabCount !== this.model.tabs.length) {
            this.tabStart = 0;
        }

        // 從外部（報表重新開啟、書籤、同步篩選器）帶進來的篩選狀態，
        // 還原成畫面上的勾選；自己剛送出的那幾次則略過，避免互相覆蓋。
        if (this.pendingSelfUpdates > 0) {
            this.pendingSelfUpdates--;
        } else if (this.settings.behavior.mode.value.value === "filter" && this.model.tabs.length > 0) {
            const jsonFilters = (options as { jsonFilters?: unknown[] }).jsonFilters as IFilter[];
            const result = this.filterEngine.restoreChecked(this.model.tabs, jsonFilters || []);
            // 認得的條件才覆寫勾選狀態；完全認不得時保留現況，
            // 避免因為欄位名稱對不上就把使用者的選擇清空。
            if (result.recognized) {
                this.checked = result.checked;
                this.selection = result.selection;
            }
        }

        // 查詢回來了，解除搜尋中的提示
        if (this.searchPending && this.query.trim() === this.appliedQuery) {
            this.searchPending = false;
        }

        this.applyTheme();
        this.render();
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        if (!this.settings) {
            this.settings = new VisualSettings();
        }
        this.settings.tabs.syncVisibility();
        return this.formattingService.buildFormattingModel(this.settings);
    }

    // ── 版面計算（響應式） ─────────────────────────────────────────

    /** 依「設定上限」與「實際寬度」算出目前一頁能放幾個頁籤 */
    private tabsPerPage(): number {
        const cap = Math.max(2, Math.min(MAX_TABS, Math.round(this.settings?.tabs.perPage.value ?? 5)));
        const navRoom = this.model.tabs.length > cap ? 56 : 0;
        const usable = Math.max(0, this.viewport.width - 16 - navRoom);
        const minTabWidth = 64;
        const fits = Math.max(1, Math.floor(usable / minTabWidth));
        return Math.max(1, Math.min(cap, fits, Math.max(1, this.model.tabs.length)));
    }

    private clampTabStart(start: number): number {
        const per = this.tabsPerPage();
        return Math.max(0, Math.min(start, Math.max(0, this.model.tabs.length - per)));
    }

    private applyTheme(): void {
        const s = this.settings;
        const style = this.root.style;
        const palette = this.host.colorPalette;
        this.isHighContrast = !!palette.isHighContrast;

        // 高對比模式一律採用作業系統指定的色彩
        const accent = this.isHighContrast ? palette.hyperlink.value : s.items.accentColor.value.value;
        const text = this.isHighContrast ? palette.foreground.value : s.items.fontColor.value.value;
        const bg = this.isHighContrast ? palette.background.value : s.items.backgroundColor.value.value;

        style.setProperty("--tts-accent", accent);
        style.setProperty("--tts-text", text);
        style.setProperty("--tts-bg", bg);
        style.setProperty("--tts-row-h", `${Math.max(18, s.items.rowHeight.value)}px`);
        style.setProperty("--tts-indent", `${Math.max(0, s.items.indent.value)}px`);
        style.setProperty("--tts-font", `${Math.max(6, s.items.fontSize.value)}px`);
        style.setProperty("--tts-title-font", `${Math.max(6, s.header.fontSize.value)}px`);
        style.setProperty("--tts-title-color", this.isHighContrast ? text : s.header.fontColor.value.value);
        style.width = `${this.viewport.width}px`;
        style.height = `${this.viewport.height}px`;
    }

    // ── 繪製 ───────────────────────────────────────────────────────

    private render(): void {
        const s = this.settings;

        // 標題
        this.headerEl.style.display = s.header.show.value ? "flex" : "none";
        this.titleEl.textContent = s.header.text.value || "";
        if (s.header.showCount.value) {
            this.countEl.style.display = "flex";
            this.renderCount();
        } else {
            this.countEl.style.display = "none";
        }

        // 搜尋
        // 說明按鈕
        this.helpBtn.style.display = s.header.show.value && s.header.showHelp.value ? "block" : "none";
        if (!s.header.showHelp.value && this.helpOpen) {
            this.toggleHelp(false);
        }

        // 工具列：這個頁籤沒有階層就沒必要顯示
        this.toolbarEl.style.display = this.hasExpandableNodes() ? "flex" : "none";

        this.searchWrapEl.style.display = s.search.show.value ? "block" : "none";
        this.renderSearchState();
        this.searchEl.placeholder = s.search.placeholder.value || "搜尋…";
        if (this.searchEl.value !== this.query) {
            this.searchEl.value = this.query;
        }

        this.renderTabs();
        this.renderList();
    }

    private renderTabs(): void {
        clearElement(this.tabsTrackEl);
        const tabs = this.model.tabs;

        if (tabs.length <= 1) {
            this.tabsEl.style.display = "none";
            return;
        }
        this.tabsEl.style.display = "flex";

        const per = this.tabsPerPage();
        this.tabStart = this.clampTabStart(this.tabStart);
        const paged = tabs.length > per;

        this.prevBtn.style.display = paged ? "block" : "none";
        this.nextBtn.style.display = paged ? "block" : "none";
        this.prevBtn.disabled = this.tabStart <= 0;
        this.nextBtn.disabled = this.tabStart + per >= tabs.length;

        for (let k = 0; k < per && this.tabStart + k < tabs.length; k++) {
            const i = this.tabStart + k;
            const tab = tabs[i];
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "tts-tab" + (i === this.activeTab ? " tts-tab-active" : "");
            btn.setAttribute("role", "tab");
            btn.setAttribute("aria-selected", String(i === this.activeTab));
            btn.title = tab.name + (tab.levelNames.length > 1 ? `（${tab.levelNames.join(" › ")}）` : "");
            btn.textContent = tab.name;
            this.paintTab(btn, tab, i === this.activeTab);

            const selectedInTab = this.countCheckedInTab(tab);
            if (selectedInTab > 0) {
                const dot = document.createElement("span");
                dot.className = "tts-tab-dot";
                btn.appendChild(dot);
            }

            btn.addEventListener("click", () => {
                if (this.activeTab === i) {
                    return;
                }
                this.activeTab = i;
                this.shiftAnchorKey = null;
                // 搜尋條件綁在原本頁籤的欄位上，換頁籤就得作廢
                this.appliedQuery = "";
                if (this.settings.behavior.resetOnTabChange.value) {
                    this.query = "";
                    this.searchEl.value = "";
                }
                // 跨頁籤篩選開啟時，自我篩選條件是「目前頁籤以外」，
                // 換頁籤就得重算並重新查詢一次。
                if (this.crossTabEnabled()) {
                    this.commit();
                } else {
                    this.render();
                }
            });
            this.tabsTrackEl.appendChild(btn);
        }
    }

    /**
     * 套用頁籤自訂背景色。
     * 未啟用、或高對比模式下不上色（高對比要照作業系統的配色走）。
     * 強調色底線與已選取圓點不受影響，仍由「項目 → 強調色彩」控制。
     */
    private paintTab(btn: HTMLElement, tab: TabModel, active: boolean): void {
        if (!this.settings.tabs.useTabColors.value || this.isHighContrast) {
            return;
        }
        const picker = this.settings.tabs.colors[tab.index];
        const base = picker?.value?.value;
        if (!base) {
            return;
        }
        // 未選取的頁籤淡一點，讓目前頁籤仍然明顯
        const bg = active ? base : shade(base, 0.45);
        btn.style.background = bg;
        const fg = contrastText(bg);
        if (fg) {
            btn.style.color = fg;
        }
        btn.classList.add("tts-tab-colored");
    }

    /**
     * 更新固定狀態列。
     *
     * 這一列永遠存在：左邊的「全選」可由格式窗格關閉，
     * 右邊的「未篩選／已選 N 項」則固定顯示在同一個位置，
     * 不會因為設定或資料狀態而跑位。
     *
     * 有搜尋字串時，全選只對搜尋結果作用。
     */
    private updateStatusBar(tab: TabModel | null, rows: Row[]): void {
        const showAll = this.settings.behavior.showSelectAll.value && !!tab && rows.length > 0;
        this.selectAllBox.style.display = showAll ? "block" : "none";
        this.selectAllLabel.style.display = showAll ? "block" : "none";

        if (!showAll) {
            this.selectAllBox.onchange = null;
            return;
        }

        const searching = this.query.trim().length > 0;

        // 搜尋中只看畫面上這些列，並排除掉「祖先也在畫面上」的重複範圍
        const scopeNodes = searching
            ? rows.filter(r => !rows.some(o => o !== r && this.isDescendantOf(r.node, o.node)))
                .map(r => r.node)
            : tab.roots;

        const leafKeys = new Set<string>();
        const collect = (n: TreeNode): void => {
            if (n.children.length === 0) {
                leafKeys.add(n.key);
                return;
            }
            n.children.forEach(collect);
        };
        scopeNodes.forEach(collect);

        let checkedCount = 0;
        leafKeys.forEach(k => { if (this.checked.has(k)) { checkedCount++; } });

        this.selectAllBox.checked = leafKeys.size > 0 && checkedCount === leafKeys.size;
        this.selectAllBox.indeterminate = checkedCount > 0 && checkedCount < leafKeys.size;
        this.selectAllLabel.textContent = searching
            ? `全選搜尋結果（${leafKeys.size}）`
            : "全選";

        // 用指派而非 addEventListener，避免每次重繪疊加處理器
        this.selectAllBox.onchange = () => {
            const on = this.selectAllBox.checked;
            scopeNodes.forEach(n => this.setNodeChecked(n, on, true));
            scopeNodes.forEach(n => this.fixAncestors(n));
            this.syncSelection(tab);
            this.commit();
        };
    }

    /**
     * 該頁籤選了幾項。用 selection 而非樹，因為非目前頁籤的項目
     * 可能已被自我篩選濾掉、不在樹上。
     */
    private countCheckedInTab(tab: TabModel): number {
        return this.selection.get(tab.index)?.size ?? 0;
    }

    private crossTabEnabled(): boolean {
        return this.settings.behavior.mode.value.value === "filter"
            && this.settings.behavior.crossTab.value;
    }

    private renderList(): void {
        clearElement(this.listEl);
        const tab = this.model.tabs[this.activeTab];

        if (!tab) {
            this.emptyEl.style.display = "block";
            this.emptyEl.textContent = this.model.isEmpty
                ? "把欄位拖入右側「頁籤 1~n 欄位」，同一個頁籤放多個欄位即形成階層。\n"
                  + "建議在「篩選依據量值」放一個量值，清單才會隨其他篩選器縮減。\n"
                  + "注意：所有頁籤的欄位必須來自模型中彼此有關聯的資料表。"
                : "此頁籤沒有資料。";
            this.updateStatusBar(null, []);
            return;
        }

        const rows = this.buildRows(tab);
        if (rows.length === 0) {
            this.emptyEl.style.display = "block";
            this.emptyEl.textContent = this.query
                ? `沒有符合「${this.query}」的項目。`
                : "此頁籤沒有資料。";
            this.updateStatusBar(tab, []);
            return;
        }
        this.emptyEl.style.display = "none";
        this.visibleRows = rows;

        this.updateStatusBar(tab, rows);

        const frag = document.createDocumentFragment();
        rows.forEach((row, index) => frag.appendChild(this.buildRowElement(tab, row, index)));
        this.listEl.appendChild(frag);
    }

    /** 依展開狀態 + 搜尋字串攤平成畫面列 */
    private buildRows(tab: TabModel): Row[] {
        const q = this.query.trim().toLowerCase();
        const searchAll = this.settings.search.searchAllLevels.value;
        const rows: Row[] = [];

        const matches = (node: TreeNode): boolean => {
            if (!q) {
                return true;
            }
            return searchAll
                ? node.searchText.indexOf(q) >= 0
                : node.label.toLowerCase().indexOf(q) >= 0;
        };

        const walk = (node: TreeNode, depth: number): void => {
            const selfHit = !q || node.label.toLowerCase().indexOf(q) >= 0;
            const subtreeHit = matches(node);
            if (q && !subtreeHit) {
                return;
            }

            const state = this.nodeCheckState(node);
            const hasChildren = node.children.length > 0;
            // 搜尋時自動展開命中的分支
            const expanded = hasChildren && (
                q ? true : (this.expanded.has(node.key) || this.defaultExpanded())
            );

            rows.push({
                node,
                depth,
                hasChildren,
                expanded,
                checked: state === "all",
                indeterminate: state === "some"
            });

            if (expanded) {
                node.children.forEach(child => {
                    // 父層自己命中時，子層全部顯示
                    if (selfHit && q) {
                        this.pushSubtree(child, depth + 1, rows);
                    } else {
                        walk(child, depth + 1);
                    }
                });
            }
        };

        tab.roots.forEach(r => walk(r, 0));
        return rows;
    }

    private pushSubtree(node: TreeNode, depth: number, rows: Row[]): void {
        const state = this.nodeCheckState(node);
        const hasChildren = node.children.length > 0;
        rows.push({
            node,
            depth,
            hasChildren,
            expanded: hasChildren,
            checked: state === "all",
            indeterminate: state === "some"
        });
        if (hasChildren) {
            node.children.forEach(c => this.pushSubtree(c, depth + 1, rows));
        }
    }

    /** 節點的三態：all / some / none */
    private nodeCheckState(node: TreeNode): "all" | "some" | "none" {
        if (this.checked.has(node.key)) {
            return "all";
        }
        if (node.children.length === 0) {
            return "none";
        }
        let all = true;
        let any = false;
        for (const child of node.children) {
            const s = this.nodeCheckState(child);
            if (s !== "all") {
                all = false;
            }
            if (s !== "none") {
                any = true;
            }
            if (!all && any) {
                break;
            }
        }
        return all ? "all" : (any ? "some" : "none");
    }

    private buildRowElement(tab: TabModel, row: Row, index: number): HTMLElement {
        const el = document.createElement("div");
        el.className = "tts-row"
            + (row.node.key === this.shiftAnchorKey ? " tts-row-anchor" : "");
        el.setAttribute("role", "treeitem");
        el.style.paddingLeft = `calc(${row.depth} * var(--tts-indent))`;
        if (row.node.key === this.shiftAnchorKey) {
            el.title = "區間起點 —— 按住 Shift 點另一列即可框選中間全部";
        }

        const box = document.createElement("input");
        box.type = "checkbox";
        box.className = "tts-check";
        box.checked = row.checked;
        box.indeterminate = row.indeterminate;
        box.setAttribute("aria-label", row.node.label);
        // click 早於 change，在這裡把 Shift 狀態記下來給 change 用
        box.addEventListener("click", (ev: MouseEvent) => {
            this.shiftPending = ev.shiftKey;
        });
        box.addEventListener("change", () => this.handleRowToggle(tab, row, index, box.checked));
        el.appendChild(box);

        if (row.hasChildren) {
            const caret = document.createElement("button");
            caret.type = "button";
            caret.className = "tts-caret" + (row.expanded ? " tts-caret-open" : "");
            caret.setAttribute("aria-label", row.expanded ? "折疊" : "展開");
            caret.textContent = "▶";
            caret.addEventListener("click", (ev) => {
                ev.stopPropagation();
                this.toggleExpand(row.node);
            });
            el.appendChild(caret);

            // 有下層時，整列按右鍵＝展開/折疊，不叫出 Power BI 的內容功能表
            el.classList.add("tts-row-expandable");
            el.addEventListener("contextmenu", (ev: MouseEvent) => {
                ev.preventDefault();
                ev.stopPropagation();
                (ev as MouseEvent & { ttsHandled?: boolean }).ttsHandled = true;
                this.toggleExpand(row.node);
            });
        } else {
            const spacer = document.createElement("span");
            spacer.className = "tts-caret-spacer";
            el.appendChild(spacer);
        }

        const label = document.createElement("div");
        label.className = "tts-label";
        label.textContent = row.node.label;
        label.title = `${tab.levelNames[row.node.level] ?? ""}：${row.node.label}`;
        label.addEventListener("click", (ev: MouseEvent) => {
            this.shiftPending = ev.shiftKey;
            this.handleRowToggle(tab, row, index, !row.checked);
        });
        el.appendChild(label);

        if (this.settings.items.showCounts.value && row.hasChildren) {
            const count = document.createElement("span");
            count.className = "tts-rowcount";
            count.textContent = String(row.node.leafCount);
            el.appendChild(count);
        }

        return el;
    }

    /**
     * 展開/折疊某個節點。只有真的有下一層的節點才會被呼叫到
     * （呼叫端已用 row.hasChildren 判斷過）。
     */
    // ── 伺服器端搜尋 ───────────────────────────────────────────────

    private serverSearchEnabled(): boolean {
        return this.settings.search.show.value
            && this.settings.search.serverSide.value
            && this.settings.behavior.mode.value.value === "filter";
    }

    /**
     * 打字時排程一次查詢。每按一個鍵就查一次太貴，
     * 所以等使用者停手 400ms 才送出。
     */
    private scheduleServerSearch(): void {
        if (!this.serverSearchEnabled()) {
            return;
        }
        if (this.searchTimer !== null) {
            clearTimeout(this.searchTimer);
        }
        this.searchPending = this.query.trim() !== this.appliedQuery;
        this.renderSearchState();

        this.searchTimer = window.setTimeout(() => {
            this.searchTimer = null;
            const next = this.query.trim();
            if (next === this.appliedQuery) {
                this.searchPending = false;
                this.renderSearchState();
                return;
            }
            this.appliedQuery = next;
            this.commit();
        }, 400);
    }

    /**
     * 目前搜尋字串要送出的查詢條件。
     * 只有伺服器端搜尋開啟、而且該頁籤有欄位時才產生。
     */
    private currentSearchFilter(): IFilter | null {
        if (!this.serverSearchEnabled() || !this.appliedQuery) {
            return null;
        }
        const tab = this.model.tabs[this.activeTab];
        if (!tab || tab.targets.length === 0) {
            return null;
        }
        // 設定值 1 = 最上層；0 或超出範圍就用最底層
        const raw = Math.round(this.settings.search.matchLevel.value ?? 0);
        const idx = raw >= 1 && raw <= tab.targets.length
            ? raw - 1
            : tab.targets.length - 1;
        return FilterEngine.buildSearchFilter(tab.targets[idx], this.appliedQuery);
    }

    private renderSearchState(): void {
        this.searchEl.classList.toggle("tts-search-pending", this.searchPending);
    }

    // ── 說明面板 ───────────────────────────────────────────────────

    private toggleHelp(force?: boolean): void {
        this.helpOpen = force !== undefined ? force : !this.helpOpen;
        this.helpPanelEl.style.display = this.helpOpen ? "flex" : "none";
        this.helpBtn.classList.toggle("tts-help-active", this.helpOpen);
        if (this.helpOpen) {
            this.renderHelp();
        }
    }

    private renderHelp(): void {
        clearElement(this.helpPanelEl);

        const head = document.createElement("div");
        head.className = "tts-helphead";
        const heading = document.createElement("div");
        heading.className = "tts-helptitle";
        heading.textContent = "操作說明";
        const close = document.createElement("button");
        close.type = "button";
        close.className = "tts-helpclose";
        close.textContent = "✕";
        close.setAttribute("aria-label", "關閉");
        close.addEventListener("click", () => this.toggleHelp(false));
        head.appendChild(heading);
        head.appendChild(close);
        this.helpPanelEl.appendChild(head);

        const body = document.createElement("div");
        body.className = "tts-helpbody";
        for (const section of HELP_SECTIONS) {
            const h = document.createElement("div");
            h.className = "tts-helpsection";
            h.textContent = section.title;
            body.appendChild(h);

            const ul = document.createElement("ul");
            ul.className = "tts-helplist";
            for (const item of section.items) {
                const li = document.createElement("li");
                li.textContent = item;
                ul.appendChild(li);
            }
            body.appendChild(ul);
        }

        // 選填的外部文件連結。一定要走 launchUrl，
        // 自訂視覺是跑在沙箱 iframe 裡，直接開新分頁會被擋掉。
        const url = (this.settings.header.helpUrl.value || "").trim();
        if (/^https?:\/\//i.test(url)) {
            const link = document.createElement("button");
            link.type = "button";
            link.className = "tts-helplink";
            link.textContent = "開啟完整說明文件 ↗";
            link.addEventListener("click", () => this.host.launchUrl(url));
            body.appendChild(link);
        }

        this.helpPanelEl.appendChild(body);
    }

    /** 沒有個別展開狀態時的預設：使用者按過工具列就依工具列，否則跟隨格式窗格 */
    private defaultExpanded(): boolean {
        return this.expandAllOverride !== null
            ? this.expandAllOverride
            : this.settings.behavior.expandAll.value;
    }

    /** 工具列的「全部展開 / 全部折疊」，只作用在目前顯示的頁籤 */
    private setAllExpanded(expand: boolean): void {
        const tab = this.model.tabs[this.activeTab];
        if (!tab) {
            return;
        }
        tab.nodeMap.forEach((node, key) => {
            if (node.children.length === 0) {
                return;
            }
            if (expand) {
                this.expanded.add(key);
            } else {
                this.expanded.delete(key);
            }
        });
        // 覆寫格式窗格的「預設全部展開」，不然折疊會被它蓋回去
        this.expandAllOverride = expand;
        this.renderList();
    }

    /** 目前這個頁籤有沒有可以展開的節點 */
    private hasExpandableNodes(): boolean {
        const tab = this.model.tabs[this.activeTab];
        if (!tab) {
            return false;
        }
        for (const node of tab.roots) {
            if (node.children.length > 0) {
                return true;
            }
        }
        return false;
    }

    private toggleExpand(node: TreeNode): void {
        if (node.children.length === 0) {
            return;
        }
        if (this.expanded.has(node.key)) {
            this.expanded.delete(node.key);
        } else {
            this.expanded.add(node.key);
        }
        this.renderList();
    }

    // ── 勾選邏輯 ───────────────────────────────────────────────────

    private toggleCheck(tab: TabModel, node: TreeNode, on: boolean): void {
        const single = this.settings.behavior.singleSelect.value;
        const cascade = this.settings.behavior.selectChildren.value;

        if (single) {
            this.checked.clear();
            this.selection.clear();
            if (on) {
                this.checked.add(node.key);
            }
            this.syncSelection(tab);
            this.commit();
            return;
        }

        this.setNodeChecked(node, on, cascade);
        this.fixAncestors(node);

        this.syncSelection(tab);
        this.commit();
    }

    /**
     * 點一列的統一入口。
     *
     * 一般點擊 = 單列切換，不影響 Shift 區間的起點。
     * Shift 點擊 = 第一次設起點（該列同時被勾選），第二次完成區間並清掉起點。
     * 所以「Shift 起點 → Shift 終點」是完整的一組動作，前面隨手勾過什麼都不影響。
     */
    private handleRowToggle(tab: TabModel, row: Row, index: number, on: boolean): void {
        const shift = this.shiftPending;
        this.shiftPending = false;

        if (!shift || this.settings.behavior.singleSelect.value) {
            this.toggleCheck(tab, row.node, on);
            return;
        }

        // 起點還在畫面上才算數（搜尋或折疊後可能already不見了）
        const from = this.shiftAnchorKey === null
            ? -1
            : this.visibleRows.findIndex(r => r.node.key === this.shiftAnchorKey);

        if (from < 0) {
            // 設定起點：該列照常勾選，並標記起來讓使用者看得到
            this.shiftAnchorKey = row.node.key;
            this.toggleCheck(tab, row.node, on);
            return;
        }

        // 完成區間，並把起點清掉，下一次 Shift 重新開始
        this.shiftAnchorKey = null;
        this.applyRange(tab, from, index, on);
    }

    /** 設定單一節點（含子孫）的勾選狀態，不處理父層與送出 */
    private setNodeChecked(node: TreeNode, on: boolean, cascade: boolean): void {
        if (on) {
            this.checked.add(node.key);
        } else {
            this.checked.delete(node.key);
        }
        if (cascade) {
            node.children.forEach(c => this.setNodeChecked(c, on, cascade));
        }
    }

    /** 往上修正父層：子項全選 → 父項也視為勾選；否則取消父項 */
    private fixAncestors(node: TreeNode): void {
        let parent = node.parent;
        while (parent) {
            const allChecked = parent.children.every(c => this.checked.has(c.key));
            if (allChecked) {
                this.checked.add(parent.key);
            } else {
                this.checked.delete(parent.key);
            }
            parent = parent.parent;
        }
    }

    /**
     * Shift 區間勾選：把畫面上 from~to 之間的列全部設成同一個狀態。
     * 索引是「目前顯示的列」的索引，所以折疊起來的子項不會被掃到 ——
     * 這跟使用者看到的範圍一致。
     */
    private applyRange(tab: TabModel, from: number, to: number, on: boolean): void {
        const cascade = this.settings.behavior.selectChildren.value;
        const lo = Math.min(from, to);
        const hi = Math.max(from, to);
        const touched: TreeNode[] = [];

        for (let i = lo; i <= hi && i < this.visibleRows.length; i++) {
            const node = this.visibleRows[i].node;
            // 祖先已經在範圍內且會連動時就跳過，避免重複套用
            if (cascade && touched.some(t => this.isDescendantOf(node, t))) {
                continue;
            }
            this.setNodeChecked(node, on, cascade);
            touched.push(node);
        }
        touched.forEach(n => this.fixAncestors(n));

        this.syncSelection(tab);
        this.commit();
    }

    private isDescendantOf(node: TreeNode, ancestor: TreeNode): boolean {
        let p = node.parent;
        while (p) {
            if (p === ancestor) {
                return true;
            }
            p = p.parent;
        }
        return false;
    }

    /**
     * 把目前頁籤的勾選狀態固化成「葉層路徑」存進 selection。
     *
     * 伺服器端搜尋開啟時，樹不一定完整 —— 搜尋條件是用 selfFilter 送回查詢的，
     * 不符合的列會直接從 dataView 消失，不只是畫面上被濾掉。若整批覆蓋
     * selection，前一次搜尋勾選、這次搜尋看不到的項目就會被當成沒勾而遺失，
     * 造成畫面顯示的勾選和實際送出的篩選條件分岔。所以這裡改成合併：
     * 目前樹上看得到的葉節點依 checked 重新計算，看不到的舊選項原樣保留。
     */
    private syncSelection(tab: TabModel): void {
        const fresh = collectSelection(tab, this.checked);
        const visible = new Set(tab.leaves.map(l => pathSignature(l.path)));
        const merged: TabSelection = new Map(fresh);
        this.selection.get(tab.index)?.forEach((path, sig) => {
            if (!visible.has(sig) && !merged.has(sig)) {
                merged.set(sig, path);
            }
        });
        if (merged.size === 0) {
            this.selection.delete(tab.index);
        } else {
            this.selection.set(tab.index, merged);
        }
    }

    /**
     * 計數與清除。
     *
     * 主要文字是**目前頁籤**的勾選數，點了只清這個頁籤 ——
     * 它就在「全選」旁邊，而全選也是針對目前頁籤，兩者的範圍要一致。
     * 其他頁籤還有勾選時，右邊會多出「清除全部」，全域清除仍然做得到。
     */
    private renderCount(): void {
        clearElement(this.countEl);

        const tab = this.model.tabs[this.activeTab];
        const here = tab ? (this.selection.get(tab.index)?.size ?? 0) : 0;
        let total = 0;
        this.selection.forEach(m => { total += m.size; });

        const main = document.createElement("span");
        main.textContent = here > 0 ? `已選 ${here} 項 · 清除` : "未篩選";
        main.className = here > 0 ? "tts-count-link" : "";
        if (here > 0 && tab) {
            main.title = `清除「${tab.name}」的勾選`;
            main.addEventListener("click", ev => {
                ev.stopPropagation();
                this.clearTab(tab);
            });
        }
        this.countEl.appendChild(main);

        // 只有在別的頁籤也有勾選時才出現，平常不佔位
        if (total > here) {
            const all = document.createElement("span");
            all.className = "tts-count-link tts-count-all";
            all.textContent = `全部 ${total} · 清除`;
            all.title = "清除所有頁籤的勾選";
            all.addEventListener("click", ev => {
                ev.stopPropagation();
                this.clearAll();
            });
            this.countEl.appendChild(all);
        }
    }

    /** 只清掉某個頁籤的勾選 */
    private clearTab(tab: TabModel): void {
        let changed = false;
        tab.nodeMap.forEach((_n, key) => {
            if (this.checked.delete(key)) {
                changed = true;
            }
        });
        if (this.selection.delete(tab.index)) {
            changed = true;
        }
        if (changed) {
            this.commit();
        }
    }

    private clearAll(): void {
        if (this.checked.size === 0 && this.selection.size === 0) {
            return;
        }
        this.checked.clear();
        this.selection.clear();
        this.commit();
    }

    /** 勾選狀態變更後：套用篩選 + 重繪 */
    private commit(): void {
        // 報表設定為「不允許互動」時只更新畫面，不送出篩選
        const allowInteractions = this.host.hostCapabilities.allowInteractions !== false;
        if (!allowInteractions) {
            this.render();
            return;
        }

        const mode = this.settings.behavior.mode.value.value;
        if (mode === "select") {
            // 先清掉「篩選器」模式留下的條件，否則報表還是被舊條件篩著，
            // 交叉醒目提示看起來就像完全沒生效。
            this.pendingSelfUpdates += this.filterEngine.clearFilters();
            this.filterEngine.applySelection(this.model.tabs, this.checked);
        } else {
            this.filterEngine.clearSelection();
            const activeIndex = this.model.tabs[this.activeTab]?.index ?? -1;
            this.pendingSelfUpdates += this.filterEngine.applyFilters(
                this.model.tabs,
                this.selection,
                activeIndex,
                this.settings.behavior.crossTab.value,
                this.currentSearchFilter()
            );
        }
        this.render();
    }

    private scrollTabs(step: number): void {
        this.tabStart = this.clampTabStart(this.tabStart + step);
        this.renderTabs();
    }
}
