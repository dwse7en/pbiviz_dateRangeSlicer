/*
 * 核心视觉对象逻辑 (Core Visual Logic)
 * 包含多语言支持 (Multi-language) 与高级筛选器 (Advanced Filter) 触发。
 */
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";

import { AdvancedFilter, IFilterColumnTarget } from "powerbi-models";

export class Visual implements IVisual {
    private target: HTMLElement;
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // UI Elements
    private container: HTMLDivElement;
    private startDateInput: HTMLInputElement;
    private endDateInput: HTMLInputElement;
    private headerInput: HTMLInputElement;
    // references to icon buttons so we can style them later
    private clearIconButton: HTMLButtonElement | null = null;
    private resetIconButton: HTMLButtonElement | null = null;

    // colors cached from formatting
    private _iconClearColor: string | null = null;
    private _iconResetColor: string | null = null;

    private readonly CLEAR_ICON_SVG = `
    <svg width="16" height="16" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M44.7818 24.1702L31.918 7.09935L14.1348 20.5L27.5 37L30.8556 34.6643L44.7818 24.1702Z" 
        stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>
    <path d="M27.4998 37L23.6613 40.0748L13.0978 40.074L10.4973 36.6231L4.06543 28.0876L14.4998 20.2248" 
        stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>
    <path d="M13.2056 40.072L44.5653 40.072" 
        stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
    </svg>`;

    private readonly RESET_ICON_SVG = ` 
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M21 21l-6 -6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
    <path d="M3.268 12.043a7.017 7.017 0 0 0 6.634 4.957a7.012 7.012 0 0 0 7.043 -6.131a7 7 0 0 0 -5.314 -7.672a7.021 7.021 0 0 0 -8.241 4.403" 
        stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
    <path d="M3 4v4h4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>`;
    
    // State
    private currentTarget: IFilterColumnTarget | null = null;
    private latestDataView: powerbi.DataView | null = null; // 保存最近一次的数据视图，用于默认度量等
    private initialMin: string = "";
    private initialMax: string = "";
    private isLocaleZH: boolean;
    // 当用户点击清除时，避免 update() 自动重新应用拉入的度量默认范围
    private suppressDefaultRangeApply: boolean = false;
    // 保存上一次应用的度量值，用于检测度量值是否改变
    private lastAppliedStartMeasure: string = "";
    private lastAppliedEndMeasure: string = "";
    // 标记用户是否主动清除了过滤器
    private isUserCleared: boolean = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.target = options.element;
        this.formattingSettingsService = new FormattingSettingsService();
        this.isLocaleZH = this.host.locale.indexOf("zh") === 0;

        this.initUI();
    }


    private initUI() {
        this.container = document.createElement("div");
        this.container.className = "date-slicer-container";

        // Header row: editable title and hover-revealed icons (clear/reset)
        const headerRow = document.createElement("div");
        headerRow.className = "header-row";

        this.headerInput = document.createElement("input");
        this.headerInput.type = "text";
        this.headerInput.className = "header-input";
        this.headerInput.readOnly = true; // editing moved to formatting pane
        this.headerInput.placeholder = this.isLocaleZH ? "字段名" : "Field";

        const iconsContainer = document.createElement("div");
        iconsContainer.className = "icons";
        this.clearIconButton = this.createIconButton(this.CLEAR_ICON_SVG, () => this.clearFilter(), this.isLocaleZH ? "清除" : "Clear");
        this.resetIconButton = this.createIconButton(this.RESET_ICON_SVG, () => this.applyDefaultMeasures(), this.isLocaleZH ? "重置" : "Reset");
        iconsContainer.append(this.clearIconButton, this.resetIconButton);

        headerRow.append(this.headerInput, iconsContainer);
        this.container.append(headerRow);
        // 1. Date Inputs Panel (双日期输入框)
        const inputsPanel = document.createElement("div");
        inputsPanel.className = "inputs-panel";

        const startGroup = document.createElement("div");
        startGroup.className = "input-group";

        this.startDateInput = document.createElement("input");
        this.startDateInput.type = "date";
        // remember which input was last edited so we can auto-correct invalid ranges
        this.startDateInput.addEventListener("change", (e) => this.onDateChange(e as Event));
        startGroup.append(this.startDateInput);

        const endGroup = document.createElement("div");
        endGroup.className = "input-group";
        this.endDateInput = document.createElement("input");
        this.endDateInput.type = "date";
        this.endDateInput.addEventListener("change", (e) => this.onDateChange(e as Event));
        endGroup.append(this.endDateInput);

        inputsPanel.append(startGroup, endGroup);
        
        // append elements into main container
        this.container.append(inputsPanel);
        this.target.appendChild(this.container);
    }

    private createButton(text: string, onClick: () => void, extraClass?: string): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.innerText = text;
        if (extraClass) btn.classList.add(extraClass);
        btn.addEventListener("click", onClick);
        return btn;
    }

    private createIconButton(icon: string, onClick: () => void, title?: string): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.className = "icon-btn";
        btn.innerHTML = icon;
        if (title) btn.title = title;
        btn.addEventListener("click", onClick);
        // colors will be applied in applyFormatting()
        return btn;
    }

    public update(options: VisualUpdateOptions) {
        const dataView = options.dataViews && options.dataViews[0];
        this.latestDataView = dataView || null;
        // 提取格式化设置 (Formatting Pane)
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        this.applyFormatting();

        if (!dataView || !dataView.categorical || !dataView.categorical.categories) return;

        const category = dataView.categorical.categories[0];
        
        // 提取数据视图映射 (Data View Mappings) 中的目标字段，作为过滤目标
        this.currentTarget = {
            table: category.source.queryName.substr(0, category.source.queryName.indexOf('.')),
            column: category.source.displayName
        };

        // header text is obtained from formatting settings; default to field name if empty
        if (this.headerInput) {
            const textSetting = this.formattingSettings?.headerTextCard?.headerText?.value;
            this.headerInput.value = textSetting && textSetting.length ? textSetting : category.source.displayName;
        }

        // 默认值：当没有用户手动选择时，使用类别中的最小/最大日期，并设置控件范围
        if (category.values && category.values.length) {
            const dates = category.values
                .map(v => new Date(v as any).getTime())
                .filter(t => !isNaN(t));
            if (dates.length) {
                const min = new Date(Math.min(...dates));
                const max = new Date(Math.max(...dates));
                const minStr = this.formatDate(min);
                const maxStr = this.formatDate(max);
                this.initialMin = minStr;
                this.initialMax = maxStr;
                this.startDateInput.min = minStr;
                this.startDateInput.max = maxStr;
                this.endDateInput.min = minStr;
                this.endDateInput.max = maxStr;
            }
        }

        // 仅在当前输入尚未设置（首次渲染或返回页面）且未被用户清除时，尝试从 options 恢复筛选
        let hasRestoredFromFilter = false;
        if (!this.startDateInput.value && !this.endDateInput.value && !this.isUserCleared) {
            hasRestoredFromFilter = this.restoreFilterFromOptions(options);
            if (hasRestoredFromFilter) {
                // 恢复成功后清理标记并返回，避免自动应用度量覆盖恢复的筛选
                this.isUserCleared = false;
                this.suppressDefaultRangeApply = true;
                return;
            }
        }

        // 检查度量值是否存在及是否发生变化
        const hasDefaultMeasures = dataView.categorical.values && dataView.categorical.values.length >= 2;

        if (hasDefaultMeasures && !this.isUserCleared && !this.suppressDefaultRangeApply) {
            const startVal = dataView.categorical.values[0].values[0] as any;
            const endVal = dataView.categorical.values[1].values[0] as any;
            const currentStartMeasure = startVal != null && startVal !== true ? this.formatDate(startVal) : "";
            const currentEndMeasure = endVal != null && endVal !== true ? this.formatDate(endVal) : "";
            
            // 检测度量值是否改变
            if (currentStartMeasure !== this.lastAppliedStartMeasure || currentEndMeasure !== this.lastAppliedEndMeasure) {
                this.lastAppliedStartMeasure = currentStartMeasure;
                this.lastAppliedEndMeasure = currentEndMeasure;
                // 如果度量值改变，自动应用新的度量值
                if (currentStartMeasure) {
                    this.startDateInput.value = currentStartMeasure;
                }
                if (currentEndMeasure) {
                    this.endDateInput.value = currentEndMeasure;
                }
                this.validateDates();
                this.applyFilter();
                return;
            }
        }

        // 没有可恢复的筛选，也没有新的度量变化，需要根据当前状态填充或保持范围
        if (!this.suppressDefaultRangeApply && hasDefaultMeasures) {
            this.applyDefaultMeasures();
        } else if (this.suppressDefaultRangeApply) {
            // clear 后保持输入框为初始 min/max 值
            if (this.initialMin) {
                this.startDateInput.min = this.initialMin;
                this.endDateInput.min = this.initialMin;
            }
            if (this.initialMax) {
                this.startDateInput.max = this.initialMax;
                this.endDateInput.max = this.initialMax;
            }
            // 不调用 applyFilter()，以保持已清除的无筛选状态
        } else {
            // 没有度量值也未被抑制：使用初始 min/max 填充并应用过滤
            if (!this.startDateInput.value && this.initialMin) {
                this.startDateInput.value = this.initialMin;
            }
            if (!this.endDateInput.value && this.initialMax) {
                this.endDateInput.value = this.initialMax;
            }
            this.validateDates();
            this.applyFilter();
        }
    }

    private hexToRgb(hex: string): { r: number, g: number, b: number } | null {
        const shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
        const fullHex = hex.replace(shorthandRegex, (m, r, g, b) => r + r + g + g + b + b);
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(fullHex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
        } : null;
    }

    private applyFormatting() {
        if (!this.formattingSettings) return;
        const inputSettings = this.formattingSettings.dateInputsCard;

        // header formatting / text value
        const headerSettings = (this.formattingSettings as any).headerTextCard;
        if (headerSettings && this.headerInput) {
            // text value may change via settings; update immediately
            const textSetting = headerSettings.headerText && headerSettings.headerText.value;
            if (textSetting && textSetting.length) {
                this.headerInput.value = textSetting;
            }

            this.headerInput.style.fontFamily = headerSettings.fontFamily.value;
            this.headerInput.style.color = headerSettings.fontColor.value.value;
            this.headerInput.style.fontSize = `${headerSettings.fontSize.value}px`;
            this.headerInput.style.fontWeight = headerSettings.bold.value ? "bold" : "normal";
            this.headerInput.style.fontStyle = headerSettings.italic.value ? "italic" : "normal";
            this.headerInput.style.textDecoration = headerSettings.underline.value ? "underline" : "none";
            // header background with transparency on header wrapper
            const headerBgHex = headerSettings.headerBackgroundColor.value.value;
            const headerBgTrans = headerSettings.headerBackgroundTransparency.value;
            const headerRow = this.headerInput.parentElement as HTMLElement;
            const headerRgb = this.hexToRgb(headerBgHex);
            if (headerRow) {
                if (headerRgb) {
                    const alphaBg = (100 - headerBgTrans) / 100;
                    headerRow.style.backgroundColor = `rgba(${headerRgb.r}, ${headerRgb.g}, ${headerRgb.b}, ${alphaBg})`;
                } else {
                    headerRow.style.backgroundColor = headerBgHex;
                }
                // apply top/bottom margins (as padding) from formatting
                try {
                    const top = headerSettings.headerMarginTop && headerSettings.headerMarginTop.value != null ? headerSettings.headerMarginTop.value : 6;
                    const bottom = headerSettings.headerMarginBottom && headerSettings.headerMarginBottom.value != null ? headerSettings.headerMarginBottom.value : 6;
                    headerRow.style.paddingTop = `${top}px`;
                    headerRow.style.paddingBottom = `${bottom}px`;
                } catch (e) {
                    // ignore if settings not present
                }
            }
        }
        // icon colors
        const iconSettings = (this.formattingSettings as any).headerIconCard;
        if (iconSettings) {
            this._iconClearColor = iconSettings.clearIconColor.value.value;
            this._iconResetColor = iconSettings.resetIconColor.value.value;
        }
        // update icon button styles if already created
        if (this.clearIconButton && this._iconClearColor) {
            this.clearIconButton.style.color = this._iconClearColor;
        }
        if (this.resetIconButton && this._iconResetColor) {
            this.resetIconButton.style.color = this._iconResetColor;
        }

        this.startDateInput.style.fontFamily = inputSettings.fontFamily.value;
        this.startDateInput.style.color = inputSettings.fontColor.value.value;
        this.startDateInput.style.fontSize = `${inputSettings.fontSize.value}px`;
        this.endDateInput.style.fontFamily = inputSettings.fontFamily.value;
        this.endDateInput.style.color = inputSettings.fontColor.value.value;
        this.endDateInput.style.fontSize = `${inputSettings.fontSize.value}px`;

        // 输入框背景颜色 + 透明度
        const inputBgHex = inputSettings.inputBackgroundColor.value.value;
        const inputBgTrans = inputSettings.inputBackgroundTransparency.value;
        const inputBgRgb = this.hexToRgb(inputBgHex);
        if (inputBgRgb) {
            const alphaBg = (100 - inputBgTrans) / 100;
            const bgColor = `rgba(${inputBgRgb.r}, ${inputBgRgb.g}, ${inputBgRgb.b}, ${alphaBg})`;
            this.startDateInput.style.backgroundColor = bgColor;
            this.endDateInput.style.backgroundColor = bgColor;
        } else {
            this.startDateInput.style.backgroundColor = inputBgHex;
            this.endDateInput.style.backgroundColor = inputBgHex;
        }

        // 边框颜色
        const borderColor = inputSettings.inputBorderColor.value.value;
        this.startDateInput.style.borderColor = borderColor;
        this.endDateInput.style.borderColor = borderColor;

        // 文本样式
        this.startDateInput.style.fontWeight = inputSettings.bold.value ? "bold" : "normal";
        this.endDateInput.style.fontWeight = inputSettings.bold.value ? "bold" : "normal";
        this.startDateInput.style.fontStyle = inputSettings.italic.value ? "italic" : "normal";
        this.endDateInput.style.fontStyle = inputSettings.italic.value ? "italic" : "normal";
        this.startDateInput.style.textDecoration = inputSettings.underline.value ? "underline" : "none";
        this.endDateInput.style.textDecoration = inputSettings.underline.value ? "underline" : "none";
    }

    private applyFilter() {
        if (!this.currentTarget || (!this.startDateInput.value && !this.endDateInput.value)) return;

        // if validation error exists, do not apply
        if (this.startDateInput.classList.contains("invalid") || this.endDateInput.classList.contains("invalid")) {
            return;
        }

        const conditions: any[] = [];

        const parseLocalDate = (isoDate: string): Date | null => {
            if (!isoDate) {
                return null;
            }
            const parts = isoDate.split("-").map(p => parseInt(p, 10));
            if (parts.length !== 3 || parts.some(isNaN)) {
                return null;
            }
            const [y, m, d] = parts;
            return new Date(y, m - 1, d); // local midnight
        };

        if (this.startDateInput.value) {
            const startDate = parseLocalDate(this.startDateInput.value);
            if (startDate) {
                conditions.push({
                    operator: "GreaterThanOrEqual",
                    value: startDate // pass Date object instead of ISO string
                });
            }
        }
        if (this.endDateInput.value) {
            const endDate = parseLocalDate(this.endDateInput.value);
            if (endDate) {
                conditions.push({
                    operator: "LessThanOrEqual",
                    value: endDate
                });
            }
        }

        const filter = new AdvancedFilter(this.currentTarget, "And", conditions);
        this.host.applyJsonFilter(filter, "general", "filter", powerbi.FilterAction.merge);
    }

    private restoreFilterFromOptions(options: VisualUpdateOptions): boolean {
        // 从 options.jsonFilters 中尝试恢复过滤器状态
        if (!options.jsonFilters || !Array.isArray(options.jsonFilters) || options.jsonFilters.length === 0) {
            return false;
        }

        try {
            for (const filter of options.jsonFilters) {
                if (filter && (filter as any).conditions) {
                    const conditions = (filter as any).conditions;
                    let startDate: string | null = null;
                    let endDate: string | null = null;

                    // 从条件中提取开始日期和结束日期
                    for (const condition of conditions) {
                        if (condition.operator === "GreaterThanOrEqual" && condition.value) {
                            startDate = this.formatDate(new Date(condition.value));
                        } else if (condition.operator === "LessThanOrEqual" && condition.value) {
                            endDate = this.formatDate(new Date(condition.value));
                        }
                    }

                    // 如果成功提取到日期，设置输入框并返回 true
                    if (startDate || endDate) {
                        if (startDate) {
                            this.startDateInput.value = startDate;
                        }
                        if (endDate) {
                            this.endDateInput.value = endDate;
                        }
                        this.validateDates();
                        // 恢复过滤器时，清除用户清除标记，允许日后的度量值变化更新
                        this.isUserCleared = false;
                        return true;
                    }
                }
            }
        } catch (e) {
            console.error("Error restoring filter from options:", e);
        }

        return false;
    }

    private clearFilter() {
        this.startDateInput.value = this.initialMin;  // 清除时显示最小日期
        this.endDateInput.value = this.initialMax;    // 清除时显示最大日期
        this.startDateInput.classList.remove("invalid");
        this.endDateInput.classList.remove("invalid");
        // 标记用户已主动清除
        this.isUserCleared = true;
        this.suppressDefaultRangeApply = true;
        this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
    }

    private applyDefaultMeasures() {
        // 用户主动点击“重置”时允许应用默认度量（解除清除时的抑制）
        this.suppressDefaultRangeApply = false;
        this.isUserCleared = false;  // 清除用户清除标记

        const dv = this.latestDataView;
        if (dv && dv.categorical && dv.categorical.values) {
            const values = dv.categorical.values;
            // 约定第一个为 defaultStart, 第二个为 defaultEnd
            if (values.length >= 2) {
                const startVal = values[0].values[0] as any; // could be number/string/Date/true
                const endVal = values[1].values[0] as any;
                if (startVal != null && startVal !== true) {
                    this.startDateInput.value = this.formatDate(startVal);
                }
                if (endVal != null && endVal !== true) {
                    this.endDateInput.value = this.formatDate(endVal);
                }
                // 验证并应用过滤
                this.validateDates();
                this.applyFilter();
                return;
            }
        }

        // 没有度量值，行为同 clear：显示完整范围并移除过滤器
        this.startDateInput.value = this.initialMin;
        this.endDateInput.value = this.initialMax;
        this.startDateInput.classList.remove("invalid");
        this.endDateInput.classList.remove("invalid");
        this.isUserCleared = true;            // 保持与 clear 的一致标记
        this.suppressDefaultRangeApply = true; // 保持后续 update 不重写
        this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
    }

    // store which field triggered the change so validation can correct the other value if needed
    private lastChanged: "start" | "end" | null = null;

    private onDateChange(e: Event) {
        const target = e.target as HTMLInputElement;
        if (target === this.startDateInput) {
            this.lastChanged = "start";
        } else if (target === this.endDateInput) {
            this.lastChanged = "end";
        }

        // 用户主动修改日期后，不再允许后续 update 自动将度量值设回控件
        this.suppressDefaultRangeApply = true;

        // adjust bounds of the opposite field
        const startVal = this.startDateInput.value;
        const endVal = this.endDateInput.value;
        if (startVal) {
            this.endDateInput.min = startVal;
        } else {
            this.endDateInput.min = this.initialMin;
        }
        if (endVal) {
            this.startDateInput.max = endVal;
        } else {
            this.startDateInput.max = this.initialMax;
        }

        this.validateDates();
        this.applyFilter();
    }

    private validateDates() {
        let start = this.startDateInput.value ? new Date(this.startDateInput.value) : null;
        let end = this.endDateInput.value ? new Date(this.endDateInput.value) : null;

        if (start && end && start > end) {
            // auto-correct based on which field changed last
            if (this.lastChanged === "start") {
                // bump end up to match start
                this.endDateInput.value = this.startDateInput.value;
            } else if (this.lastChanged === "end") {
                // pull start down to match end
                this.startDateInput.value = this.endDateInput.value;
            } else {
                // no record of which changed (e.g. initial sync) - just swap values
                this.startDateInput.value = this.formatDate(end);
                this.endDateInput.value = this.formatDate(start);
            }
            // update start/end variables after correction
            start = this.startDateInput.value ? new Date(this.startDateInput.value) : null;
            end = this.endDateInput.value ? new Date(this.endDateInput.value) : null;
        }

        // after potential adjustment, clear invalid flags
        this.startDateInput.classList.remove("invalid");
        this.endDateInput.classList.remove("invalid");
    }

    private formatDate(date: Date | string): string {
        const d = new Date(date);
        if (isNaN(d.getTime())) return "";
        const yyyy = d.getFullYear();
        const mm = (d.getMonth() + 1).toString().padStart(2, "0");
        const dd = d.getDate().toString().padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}