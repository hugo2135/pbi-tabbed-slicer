/**
 * dataModel.ts
 * ------------------------------------------------------------------
 * 把 Power BI 送進來的 categorical dataView 轉成畫面用的階層樹。
 *
 * 設計重點：
 *  - Power BI 一個視覺只跑一次查詢，所有頁籤的欄位都在同一個 dataView 裡，
 *    因此先依資料角色（tab1..tab8）把欄位分群，每一群就是一個頁籤。
 *  - 同一個頁籤內放入多個欄位＝該頁籤的階層層級（第 1 個欄位為最上層），
 *    層數不限。
 *  - 各頁籤的欄位在查詢中會交叉組合；建樹時以 key 去重，畫面呈現才會正確。
 */

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import PrimitiveValue = powerbi.PrimitiveValue;

import { IFilterColumnTarget } from "powerbi-models";

export const MAX_TABS = 20;

/** 與 capabilities.json 的 dataReductionAlgorithm.window.count 一致 */
export const ROW_LIMIT = 30000;

export interface TreeNode {
    /** 唯一鍵：tabIndex + 各層值串接 */
    key: string;
    /** 顯示文字 */
    label: string;
    /** 原始值（送進 filter 用，保留型別） */
    value: PrimitiveValue;
    /** 0 = 最上層 */
    level: number;
    children: TreeNode[];
    parent: TreeNode | null;
    /** 由根到此節點的原始值路徑，長度 = level + 1 */
    path: PrimitiveValue[];
    /** 此節點對應的 dataView 列索引（取第一筆即可） */
    rowIndex: number;
    /** 交叉醒目提示模式用 */
    selectionId: ISelectionId;
    /** 底下所有葉節點數量（自己是葉時為 1） */
    leafCount: number;
    /** 供搜尋比對用的小寫字串（含自身與所有子孫 label） */
    searchText: string;
}

export interface TabModel {
    /** 資料角色索引 0..MAX_TABS-1 */
    index: number;
    /** 角色名稱 tab1..tab8 */
    role: string;
    /** 頁籤顯示名稱 */
    name: string;
    /** 每一層對應的欄位名稱 */
    levelNames: string[];
    /** 每一層對應的 filter target（供 Advanced/Tuple Filter 使用） */
    targets: IFilterColumnTarget[];
    /** 每一層的 category 欄位（供 selection 使用） */
    categories: DataViewCategoryColumn[];
    roots: TreeNode[];
    /** key -> node，方便還原勾選狀態 */
    nodeMap: Map<string, TreeNode>;
    /** 所有葉節點 */
    leaves: TreeNode[];
}

export interface SlicerModel {
    tabs: TabModel[];
    /** 是否完全沒有繫結任何欄位 */
    isEmpty: boolean;
    /** 這次查詢回傳的列數 */
    rowCount: number;
    /** 是否已達列數上限（清單可能不完整） */
    truncated: boolean;
}

/**
 * 從 column metadata 取出 filter target。
 * 一定要用 queryName 拆出「模型中的真實表名 / 欄位名」，
 * 不能用 displayName —— 使用者在視覺上重新命名欄位後 displayName 會變，
 * 篩選就會打不到、也會讓還原比對失敗。
 */
function buildTarget(column: powerbi.DataViewMetadataColumn): IFilterColumnTarget {
    // queryName 形如 "Table.Column"、"Table.Hierarchy.Level"，
    // 表名含點號時會被引號包起來："My.Table".Column
    const queryName = column.queryName || "";
    let table = queryName;
    let columnName = column.displayName;

    if (queryName.charAt(0) === "\"") {
        const close = queryName.indexOf("\"", 1);
        if (close > 0) {
            table = queryName.substring(1, close);
            columnName = queryName.substring(close + 2) || columnName;
        }
    } else {
        const dot = queryName.indexOf(".");
        if (dot > 0) {
            table = queryName.substring(0, dot);
            columnName = queryName.substring(dot + 1) || columnName;
        }
    }

    // 階層層級 "Table.Hierarchy.Level" → 取最後一段當欄位名
    const lastDot = columnName.lastIndexOf(".");
    if (lastDot > 0) {
        columnName = columnName.substring(lastDot + 1);
    }

    return { table, column: columnName };
}

/** 找出某個 category 欄位屬於哪個頁籤角色（tab1..tab8） */
function detectRole(column: DataViewCategoryColumn): number {
    const roles = column.source?.roles;
    if (!roles) {
        return -1;
    }
    for (let i = 1; i <= MAX_TABS; i++) {
        if (roles[`tab${i}`]) {
            return i - 1;
        }
    }
    return -1;
}

function formatValue(v: PrimitiveValue): string {
    if (v === null || v === undefined || v === "") {
        return "(空白)";
    }
    if (v instanceof Date) {
        return v.toLocaleDateString();
    }
    return String(v);
}

/**
 * 建立單一頁籤的樹。
 * categories 已依「欄位放入順序」排好，第 0 個為最上層。
 */
function buildTree(
    host: IVisualHost,
    tabIndex: number,
    categories: DataViewCategoryColumn[]
): { roots: TreeNode[]; nodeMap: Map<string, TreeNode>; leaves: TreeNode[] } {
    const roots: TreeNode[] = [];
    const nodeMap = new Map<string, TreeNode>();
    const leaves: TreeNode[] = [];
    const rowCount = categories[0].values?.length ?? 0;
    const depth = categories.length;

    for (let row = 0; row < rowCount; row++) {
        let parent: TreeNode | null = null;
        let keyPrefix = `t${tabIndex}`;
        const path: PrimitiveValue[] = [];

        for (let level = 0; level < depth; level++) {
            const raw = categories[level].values[row];
            path.push(raw);
            const key = `${keyPrefix}|${level}:${String(raw)}`;
            keyPrefix = key;

            let node = nodeMap.get(key);
            if (!node) {
                // selectionId 只綁「這個節點自己那一層」的欄位。
                //
                // 曾經改成把根到本層的所有欄位串起來（連續呼叫 withCategory），
                // 想讓交叉醒目提示更精準 —— 結果 Power BI 無法把那種複合識別
                // 對應到任何資料，選取變成完全沒有反應。單一 withCategory
                // 才是 SDK 實際支援的用法。
                const builder = host.createSelectionIdBuilder()
                    .withCategory(categories[level], row);
                node = {
                    key,
                    label: formatValue(raw),
                    value: raw,
                    level,
                    children: [],
                    parent,
                    path: path.slice(),
                    rowIndex: row,
                    selectionId: builder.createSelectionId(),
                    leafCount: 0,
                    searchText: ""
                };
                nodeMap.set(key, node);
                if (parent) {
                    parent.children.push(node);
                } else {
                    roots.push(node);
                }
            }
            parent = node;
        }
    }

    // 後序走訪：計算 leafCount 與 searchText
    const finalize = (node: TreeNode): void => {
        if (node.children.length === 0) {
            node.leafCount = 1;
            node.searchText = node.label.toLowerCase();
            leaves.push(node);
            return;
        }
        let count = 0;
        let text = node.label.toLowerCase();
        for (const child of node.children) {
            finalize(child);
            count += child.leafCount;
            text += "\u0001" + child.searchText;
        }
        node.leafCount = count;
        node.searchText = text;
    };
    roots.forEach(finalize);

    return { roots, nodeMap, leaves };
}

export function transform(
    host: IVisualHost,
    dataView: DataView,
    tabNames: string[],
    autoName: boolean
): SlicerModel {
    const tabs: TabModel[] = [];
    const categories = dataView?.categorical?.categories;
    if (!categories || categories.length === 0) {
        return { tabs, isEmpty: true, rowCount: 0, truncated: false };
    }

    // Power BI 一個視覺只跑一次查詢，所有頁籤的欄位都在同一個 dataView 裡，
    // 因此先依資料角色把欄位分群，每一群就是一個頁籤的階層。
    const byRole = new Map<number, DataViewCategoryColumn[]>();
    for (const column of categories) {
        const index = detectRole(column);
        if (index < 0) {
            continue;
        }
        const list = byRole.get(index) || [];
        list.push(column);
        byRole.set(index, list);
    }

    const rowCount = categories[0].values?.length ?? 0;

    byRole.forEach((cols, index) => {
        // 只有「完全沒繫結欄位」才不算一個頁籤。
        // 資料被篩到 0 列時仍要保留這個頁籤 —— 否則它的勾選狀態會被當成
        // 失效而清掉，條件跟著消失、資料又回來，形成來回震盪。
        if (cols.length === 0) {
            return;
        }
        const levelNames = cols.map(c => c.source.displayName);
        const targets = cols.map(c => buildTarget(c.source));
        // buildTree 以 key 去重，因此各頁籤欄位交叉合併產生的重複列會自動收斂
        const { roots, nodeMap, leaves } = buildTree(host, index, cols);

        // 命名優先序：使用者自訂 > 自動取欄位名稱 > 「頁籤 N」
        const custom = (tabNames[index] || "").trim();
        const name = custom
            ? custom
            : (autoName ? (levelNames[0] || `頁籤 ${index + 1}`) : `頁籤 ${index + 1}`);

        tabs.push({
            index,
            role: `tab${index + 1}`,
            name,
            levelNames,
            targets,
            categories: cols,
            roots,
            nodeMap,
            leaves
        });
    });

    tabs.sort((a, b) => a.index - b.index);
    return {
        tabs,
        isEmpty: tabs.length === 0,
        rowCount,
        // 列數頂到上限代表資料被截斷，各頁籤的清單可能不完整
        truncated: rowCount >= ROW_LIMIT
    };
}
