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
import { FilterEngine, SelectionStore, collectSelection } from "./filterEngine";

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
    private tabsEl: HTMLElement;
    private tabsTrackEl: HTMLElement;
    private prevBtn: HTMLButtonElement;
    private nextBtn: HTMLButtonElement;
    private searchWrapEl: HTMLElement;
    private searchEl: HTMLInputElement;
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
    /** 剛剛是我們自己送出篩選，下一次 update 不要再從 jsonFilters 還原 */
    private selfApplied = false;
    private events: powerbi.extensibility.IVisualEventService;
    private isHighContrast = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingService = new FormattingSettingsService();
        this.filterEngine = new FilterEngine(this.host, this.selectionManager);
        this.events = this.host.eventService;
        this.buildDom(options.element);

        // 右鍵內容功能表
        this.root.addEventListener("contextmenu", (ev: MouseEvent) => {
            ev.preventDefault();
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
        this.countEl.addEventListener("click", () => this.clearAll());
        this.headerEl.appendChild(this.titleEl);
        this.headerEl.appendChild(this.countEl);
        this.root.appendChild(this.headerEl);

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

        // 搜尋列
        this.searchWrapEl = document.createElement("div");
        this.searchWrapEl.className = "tts-searchwrap";
        this.searchEl = document.createElement("input");
        this.searchEl.type = "search";
        this.searchEl.className = "tts-search";
        this.searchEl.addEventListener("input", () => {
            this.query = this.searchEl.value;
            this.renderList();
        });
        this.searchWrapEl.appendChild(this.searchEl);
        this.root.appendChild(this.searchWrapEl);

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
        // 還原成畫面上的勾選；自己剛送出的那一次則略過，避免互相覆蓋。
        if (this.selfApplied) {
            this.selfApplied = false;
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
            let n = 0;
            this.selection.forEach(m => { n += m.size; });
            this.countEl.style.display = "block";
            this.countEl.textContent = n > 0 ? `已選 ${n} 項 · 清除` : "未篩選";
            this.countEl.classList.toggle("tts-count-active", n > 0);
        } else {
            this.countEl.style.display = "none";
        }

        // 搜尋
        this.searchWrapEl.style.display = s.search.show.value ? "block" : "none";
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
                ? "把欄位拖入右側「頁籤 1~8 欄位」，同一個頁籤放多個欄位即形成階層。\n"
                  + "建議在「篩選依據量值」放一個量值，清單才會隨其他篩選器縮減。\n"
                  + "注意：所有頁籤的欄位必須來自模型中彼此有關聯的資料表。"
                : "此頁籤沒有資料。";
            return;
        }

        const rows = this.buildRows(tab);
        if (rows.length === 0) {
            this.emptyEl.style.display = "block";
            this.emptyEl.textContent = this.query
                ? `沒有符合「${this.query}」的項目。`
                : "此頁籤沒有資料。";
            return;
        }
        this.emptyEl.style.display = "none";

        if (this.settings.behavior.showSelectAll.value && !this.query) {
            this.listEl.appendChild(this.buildSelectAllRow(tab));
        }

        const frag = document.createDocumentFragment();
        rows.forEach(row => frag.appendChild(this.buildRowElement(tab, row)));
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
                q ? true : (this.expanded.has(node.key) || this.settings.behavior.expandAll.value)
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

    private buildSelectAllRow(tab: TabModel): HTMLElement {
        const total = tab.leaves.length;
        const selected = this.countCheckedInTab(tab);
        const row = document.createElement("div");
        row.className = "tts-row tts-row-all";

        const box = document.createElement("input");
        box.type = "checkbox";
        box.className = "tts-check";
        box.checked = selected > 0 && selected >= total;
        box.indeterminate = selected > 0 && selected < total;
        box.addEventListener("change", () => {
            if (box.checked) {
                tab.nodeMap.forEach((_n, k) => this.checked.add(k));
            } else {
                tab.nodeMap.forEach((_n, k) => this.checked.delete(k));
            }
            this.syncSelection(tab);
            this.commit();
        });

        const label = document.createElement("div");
        label.className = "tts-label tts-label-all";
        label.textContent = "全選";
        label.addEventListener("click", () => {
            box.checked = !box.checked;
            box.dispatchEvent(new Event("change"));
        });

        row.appendChild(box);
        row.appendChild(label);
        return row;
    }

    private buildRowElement(tab: TabModel, row: Row): HTMLElement {
        const el = document.createElement("div");
        el.className = "tts-row";
        el.setAttribute("role", "treeitem");
        el.style.paddingLeft = `calc(${row.depth} * var(--tts-indent))`;

        const box = document.createElement("input");
        box.type = "checkbox";
        box.className = "tts-check";
        box.checked = row.checked;
        box.indeterminate = row.indeterminate;
        box.setAttribute("aria-label", row.node.label);
        box.addEventListener("change", () => this.toggleCheck(tab, row.node, box.checked));
        el.appendChild(box);

        if (row.hasChildren) {
            const caret = document.createElement("button");
            caret.type = "button";
            caret.className = "tts-caret" + (row.expanded ? " tts-caret-open" : "");
            caret.setAttribute("aria-label", row.expanded ? "折疊" : "展開");
            caret.textContent = "▶";
            caret.addEventListener("click", (ev) => {
                ev.stopPropagation();
                if (this.expanded.has(row.node.key)) {
                    this.expanded.delete(row.node.key);
                } else {
                    this.expanded.add(row.node.key);
                }
                this.renderList();
            });
            el.appendChild(caret);
        } else {
            const spacer = document.createElement("span");
            spacer.className = "tts-caret-spacer";
            el.appendChild(spacer);
        }

        const label = document.createElement("div");
        label.className = "tts-label";
        label.textContent = row.node.label;
        label.title = `${tab.levelNames[row.node.level] ?? ""}：${row.node.label}`;
        label.addEventListener("click", () => {
            const next = !(row.checked);
            this.toggleCheck(tab, row.node, next);
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

        const apply = (n: TreeNode): void => {
            if (on) {
                this.checked.add(n.key);
            } else {
                this.checked.delete(n.key);
            }
            if (cascade) {
                n.children.forEach(apply);
            }
        };
        apply(node);

        // 往上修正父層：子項全選 → 父項也視為勾選；否則取消父項
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

        this.syncSelection(tab);
        this.commit();
    }

    /**
     * 把目前頁籤的勾選狀態固化成「葉層路徑」存進 selection。
     * 只在該頁籤正在顯示、樹是完整的時候呼叫。
     */
    private syncSelection(tab: TabModel): void {
        const sel = collectSelection(tab, this.checked);
        if (sel.size === 0) {
            this.selection.delete(tab.index);
        } else {
            this.selection.set(tab.index, sel);
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
        this.selfApplied = true;
        if (mode === "select") {
            this.filterEngine.applySelection(this.model.tabs, this.checked);
        } else {
            const activeIndex = this.model.tabs[this.activeTab]?.index ?? -1;
            this.filterEngine.applyFilters(
                this.model.tabs,
                this.selection,
                activeIndex,
                this.settings.behavior.crossTab.value
            );
        }
        this.render();
    }

    private scrollTabs(step: number): void {
        this.tabStart = this.clampTabStart(this.tabStart + step);
        this.renderTabs();
    }
}
