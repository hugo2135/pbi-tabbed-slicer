/**
 * filterEngine.ts
 * ------------------------------------------------------------------
 * 篩選條件的組裝、送出與還原。
 *
 * 兩種篩選模式：
 *  1) filter — 用 Advanced Filter API 送出真正的篩選條件。
 *     單層頁籤 → BasicFilter (In)；多層頁籤 → TupleFilter（多欄位組合）。
 *     各頁籤各出一組條件，彼此 AND（等同多個原生篩選器）。
 *  2) select — 用 SelectionManager 做交叉醒目提示。
 *
 * 兩個 filter 屬性：
 *  - general.filter     → 送給「其他視覺」的條件（全部頁籤）
 *  - general.selfFilter → 套回「本視覺自己」的條件（排除目前頁籤）
 *    Power BI 預設不把視覺自己送出的篩選套回自己，所以要跨頁籤互相篩選
 *    就得靠 selfFilter。排除目前頁籤是必要的 —— 否則勾了 A 之後
 *    清單只剩 A，就再也改選不了 B。
 */

import powerbi from "powerbi-visuals-api";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import FilterAction = powerbi.FilterAction;
import PrimitiveValue = powerbi.PrimitiveValue;

import {
    IFilter,
    IBasicFilter,
    ITupleFilter,
    ITupleElementValue,
    IFilterColumnTarget,
    BasicFilter,
    TupleFilter,
    AdvancedFilter
} from "powerbi-models";

import { TabModel, TreeNode } from "./dataModel";

const FILTER_OBJECT = "general";
const FILTER_PROPERTY = "filter";
const SELF_FILTER_PROPERTY = "selfFilter";

export const PATH_SEP = "\u0001";

/**
 * 每個頁籤已勾選的「葉層路徑」。
 *
 * 為什麼不直接從樹上算？因為開啟跨頁籤篩選後，其他頁籤的項目會被
 * selfFilter 濾掉而從 dataView 消失。若每次都從樹重算，那些條件就會
 * 憑空不見 → 條件消失 → 資料回來 → 條件又出現，來回震盪。
 * 所以勾選當下（該頁籤正在顯示、樹是完整的）就把葉層路徑存起來。
 */
export type TabSelection = Map<string, PrimitiveValue[]>;
export type SelectionStore = Map<number, TabSelection>;

export interface RestoreResult {
    checked: Set<string>;
    selection: SelectionStore;
    /** 是否有任何一個篩選條件對應到本視覺的頁籤欄位 */
    recognized: boolean;
}

/** 把值正規化成 filter 可接受的基本型別 */
function normalize(v: PrimitiveValue): string | number | boolean {
    if (v === null || v === undefined) {
        return null;
    }
    if (v instanceof Date) {
        return v.toISOString();
    }
    return v as string | number | boolean;
}

export function pathSignature(path: PrimitiveValue[]): string {
    return path.map(p => String(normalize(p))).join(PATH_SEP);
}

function collectLeaves(node: TreeNode, out: TreeNode[]): void {
    if (node.children.length === 0) {
        out.push(node);
        return;
    }
    node.children.forEach(c => collectLeaves(c, out));
}

/**
 * 從樹上算出某個頁籤目前勾選的葉層路徑。
 * 只在該頁籤正在顯示（樹完整）時呼叫。
 */
export function collectSelection(tab: TabModel, checked: Set<string>): TabSelection {
    const result: TabSelection = new Map();
    const walk = (node: TreeNode): void => {
        if (checked.has(node.key)) {
            // 勾到父節點 = 底下所有葉節點都納入，維持 TupleFilter 的 arity 一致
            const leaves: TreeNode[] = [];
            collectLeaves(node, leaves);
            leaves.forEach(l => result.set(pathSignature(l.path), l.path));
            return;
        }
        node.children.forEach(walk);
    };
    tab.roots.forEach(walk);
    return result;
}

/** 由葉層路徑組出該頁籤的篩選條件 */
export function buildFilter(targets: IFilterColumnTarget[], selection: TabSelection): IFilter | null {
    if (!selection || selection.size === 0 || targets.length === 0) {
        return null;
    }
    const paths = Array.from(selection.values());

    // 單層 → BasicFilter（$schema 與 filterType 由 powerbi-models 填入）
    if (targets.length === 1) {
        return new BasicFilter(targets[0], "In", paths.map(p => normalize(p[0]))).toJSON();
    }

    const values: ITupleElementValue[][] = paths
        .filter(p => p.length === targets.length)
        .map(p => p.map(v => ({ value: normalize(v) })));

    if (values.length === 0) {
        return null;
    }
    return new TupleFilter(targets, "In", values).toJSON();
}

export class FilterEngine {
    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private lastOutCount = 0;
    private lastSelfCount = 0;

    constructor(host: IVisualHost, selectionManager: ISelectionManager) {
        this.host = host;
        this.selectionManager = selectionManager;
    }

    /**
     * filter 模式：
     *  - 對外送出全部頁籤的條件
     *  - crossTab 開啟時，另外把「目前頁籤以外」的條件套回自己
     */
    public applyFilters(
        tabs: TabModel[],
        selection: SelectionStore,
        activeTabIndex: number,
        crossTab: boolean,
        searchFilter?: IFilter | null
    ): void {
        const outward: IFilter[] = [];
        const inward: IFilter[] = [];

        for (const tab of tabs) {
            const f = buildFilter(tab.targets, selection.get(tab.index));
            if (!f) {
                continue;
            }
            outward.push(f);
            if (crossTab && tab.index !== activeTabIndex) {
                inward.push(f);
            }
        }

        // 伺服器端搜尋的條件只套回自己，不會影響其他視覺
        if (searchFilter) {
            inward.push(searchFilter);
        }

        this.push(FILTER_PROPERTY, outward, "lastOutCount");
        this.push(SELF_FILTER_PROPERTY, inward, "lastSelfCount");
    }

    /**
     * 伺服器端搜尋條件：對指定欄位做 Contains。
     *
     * 為什麼不是前端過濾？前端只看得到已經抓進來的那 30,000 列，
     * 超出上限的項目根本不在記憶體裡，怎麼搜都搜不到。
     * 把條件送進查詢，讓模型端去比對，才搜得到全部。
     *
     * 限制：AdvancedFilter 的一組條件只能指向**一個欄位**，
     * 所以多層頁籤沒辦法一次比對所有層級，要指定用哪一層。
     */
    public static buildSearchFilter(target: IFilterColumnTarget, query: string): IFilter | null {
        const text = (query || "").trim();
        if (!text || !target || !target.table || !target.column) {
            return null;
        }
        return new AdvancedFilter(target, "And", {
            operator: "Contains",
            value: text
        }).toJSON();
    }

    private push(property: string, filters: IFilter[], counter: "lastOutCount" | "lastSelfCount"): void {
        const previous = this[counter];

        if (filters.length === 0) {
            if (previous !== 0) {
                this.host.applyJsonFilter(null, FILTER_OBJECT, property, FilterAction.remove);
            }
            this[counter] = 0;
            return;
        }

        // 條件數變少代表有頁籤被清空，merge 不會移除舊條件，先整個清掉
        if (filters.length < previous) {
            this.host.applyJsonFilter(null, FILTER_OBJECT, property, FilterAction.remove);
        }

        this.host.applyJsonFilter(filters, FILTER_OBJECT, property, FilterAction.merge);
        this[counter] = filters.length;
    }

    /** select 模式：交叉醒目提示 */
    public applySelection(tabs: TabModel[], checked: Set<string>): void {
        const ids: ISelectionId[] = [];
        for (const tab of tabs) {
            const walk = (node: TreeNode): void => {
                if (checked.has(node.key)) {
                    ids.push(node.selectionId);
                    return;
                }
                node.children.forEach(walk);
            };
            tab.roots.forEach(walk);
        }

        if (ids.length === 0) {
            this.selectionManager.clear();
            return;
        }
        // 第二個參數是 multiSelect：false 代表「用這組取代目前選取」，
        // 傳陣列進去仍然是一次選取多個。
        this.selectionManager.select(ids, false);
    }

    /**
     * 只清掉送出去的篩選條件（不動 SelectionManager）。
     * 從「篩選器」切到「選取」模式時一定要呼叫 —— 否則舊條件會留在報表上，
     * 報表看起來還是被篩過的，交叉醒目提示的效果就被蓋住、像是沒生效。
     */
    public clearFilters(): void {
        if (this.lastOutCount !== 0) {
            this.host.applyJsonFilter(null, FILTER_OBJECT, FILTER_PROPERTY, FilterAction.remove);
            this.lastOutCount = 0;
        }
        if (this.lastSelfCount !== 0) {
            this.host.applyJsonFilter(null, FILTER_OBJECT, SELF_FILTER_PROPERTY, FilterAction.remove);
            this.lastSelfCount = 0;
        }
    }

    /** 只清掉交叉醒目提示（不動篩選條件）。從「選取」切回「篩選器」時呼叫。 */
    public clearSelection(): void {
        this.selectionManager.clear();
    }

    /**
     * 報表重開、書籤、篩選窗格、同步篩選器帶進來的條件，
     * 還原成勾選狀態與 selection 存量。
     */
    public restoreChecked(tabs: TabModel[], jsonFilters: IFilter[]): RestoreResult {
        const checked = new Set<string>();
        const selection: SelectionStore = new Map();

        // 外部把篩選清掉了 → 勾選也要跟著清掉
        if (!jsonFilters || jsonFilters.length === 0) {
            this.lastOutCount = 0;
            return { checked, selection, recognized: true };
        }

        const norm = (s: string): string => (s || "").trim().toLowerCase();
        const sameTarget = (a: IFilterColumnTarget, b: IFilterColumnTarget): boolean =>
            !!a && !!b && norm(a.table) === norm(b.table) && norm(a.column) === norm(b.column);

        let recognized = false;

        for (const raw of jsonFilters) {
            const f = raw as IBasicFilter & ITupleFilter;
            if (!f || !f.target) {
                continue;
            }
            const targets: IFilterColumnTarget[] = Array.isArray(f.target)
                ? (f.target as IFilterColumnTarget[])
                : [f.target as IFilterColumnTarget];

            const tab = tabs.find(t =>
                t.targets.length === targets.length &&
                t.targets.every((tt, i) => sameTarget(tt, targets[i]))
            );
            if (!tab) {
                continue;
            }
            recognized = true;

            // 條件裡的值本身就是葉層路徑，直接還原成 selection
            const tabSel: TabSelection = selection.get(tab.index) || new Map();
            if (targets.length === 1) {
                for (const v of (f.values || [])) {
                    tabSel.set(pathSignature([v as PrimitiveValue]), [v as PrimitiveValue]);
                }
            } else {
                for (const tupleRow of (f.values || []) as ITupleElementValue[][]) {
                    const path = tupleRow.map(c => c.value as PrimitiveValue);
                    tabSel.set(pathSignature(path), path);
                }
            }
            selection.set(tab.index, tabSel);

            // 再對照樹，把對應的節點標成勾選（畫面用）
            for (const leaf of tab.leaves) {
                if (tabSel.has(pathSignature(leaf.path))) {
                    checked.add(leaf.key);
                }
            }
        }

        this.markParentsChecked(tabs, checked);
        this.lastOutCount = jsonFilters.length;
        return { checked, selection, recognized };
    }

    /** 子項全選時父項也標記為勾選，維持三態顯示一致 */
    private markParentsChecked(tabs: TabModel[], checked: Set<string>): void {
        for (const tab of tabs) {
            const fix = (node: TreeNode): boolean => {
                if (node.children.length === 0) {
                    return checked.has(node.key);
                }
                const all = node.children.map(fix).every(Boolean);
                if (all) {
                    checked.add(node.key);
                }
                return all;
            };
            tab.roots.forEach(fix);
        }
    }

    public clearAll(): void {
        this.selectionManager.clear();
        if (this.lastOutCount !== 0) {
            this.host.applyJsonFilter(null, FILTER_OBJECT, FILTER_PROPERTY, FilterAction.remove);
            this.lastOutCount = 0;
        }
        if (this.lastSelfCount !== 0) {
            this.host.applyJsonFilter(null, FILTER_OBJECT, SELF_FILTER_PROPERTY, FilterAction.remove);
            this.lastSelfCount = 0;
        }
    }
}
