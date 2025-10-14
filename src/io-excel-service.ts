/*
// initializing new db
dataSource = {
    name: "UserTable",
    columns: [
        { name: "ID", type: "Integer", pk: true, autoIncrement: true, nullable: false },
        { name: "Name", type: "Text", pk: false, autoIncrement: false, nullable: false },
        { name: "Email", type: "Text", pk: false, autoIncrement: false, nullable: true },
    ],
    primaryKey: ["ID"],
    data: [
        [null, "John Doe", "johndoe@example.com"],
        [null, "Jane Smith", "janesmith@example.com"],
        [null, "Sam Wilson", "samwilson@example.com"],
    ]
};
await ioConnectDBService.init(dataSource);
xlService.createLinkedTable({workbook: "Book3", worksheet: "Sheet1", range: "C16", forceCreate: true}, dataSource, undefined).then(console.info);

// loading
ioConnectDBService.init({name: "UserTable",
    file: "interop.io/io.Connect Desktop/UserData/DEMO-INTEROP.IO/io.db"}).then(console.info);

// inserting
dataSource.data = [[null, "Mol Gad", "molgad@example.com"]]
ioConnectDBService.insertData(dataSource).then(console.info);

// query
await ioConnectDBService.executeQuery(dataSource, "select * from UserTable");

// create context menu for any sheet, which works for A1:B5
xlService.createContextMenu("send", ["io","actions"], {range: "A1:B5"}, console.info);

// create dynamic ribbon menu - add in a dropdown that can be executed from the ribbon
xlService.createDynamicRibbonMenu("Another", {range: "A1:B10"}, console.info);
*/

export type TableColumnOperation = "Add" | "Delete" | "Rename" | "Update";

export interface TableColumnOperationDescriptor {
    oldName: string;
    name: string;
    position: number | null;
    op: TableColumnOperation;
}

export enum XLRibbonObjectType {
    Button = "Button",
    DynamicMenu = "DynamicMenu",
    Separator = "Separator",
    Group = "Group",
    Tab = "Tab"
}

export interface XLRibbonObject {
    label?: string;
    image?: string;
    size?: string;
    tag?: string;
    callback?: SubscriptionInfo;
    type: XLRibbonObjectType;
    controls?: XLRibbonObject[];
    id?: string;
    screenTip?: string;
    superTip?: string;
}

export interface XLServiceResult {
    success?: boolean;
    message?: string;

    // Common properties
    workbook?: string;
    worksheet?: string;
    address?: string;
    subscriptionId?: string;

    // Table-related properties
    tableName?: string;
    columns?: TableColumnInfo[];
    rowsCount?: number;

    // CTP-related properties
    ctpHostId?: string;
    ctpStore?: any;

    // Menu-related properties
    menuId?: string;
    range?: RangeInfo;
    caption?: string;
    subscriptionInfo?: SubscriptionInfo;

    // File operations
    fileName?: string;

    // Window properties
    activeWindow?: string;

    // Ribbon properties
    customTabs?: XLRibbonObject[];
    customRibbonDataLocation?: string;

    // Data properties
    data?: any;
    menu?: any;
}

export interface TableColumnInfo {
    name?: string;
    address?: string;
}


export enum LifetimeType {
    None = "None",
    IOConnectInstance = "GlueInstance",
    Forever = "Forever",
    ExcelSession = "ExcelSession"
}

export interface CTPDescriptor {
    id: string;
    title: string;
    visible?: boolean;
    ui: UIDescriptor;
}

export interface UIDescriptor {
    type: UIType;
    id?: string;
    text?: string;
    callback?: CallbackInfo;
    children?: UIDescriptor[];
    horizontalAlignment?: UIHorizontalAlignment;
    verticalAlignment?: UIVerticalAlignment;
    margin?: Thickness;
    backColor?: string;
    foreColor?: string;
    isReadOnly?: boolean;
}

export type Thickness = { left: number; top: number; right: number; bottom: number };

export type UIType = "Panel" | "Label" | "TextBox" | "Button" | "ScrollBox" | "Border";

export type UIHorizontalAlignment = "Left" | "Center" | "Right" | "Stretch";
export type UIVerticalAlignment = "Top" | "Center" | "Bottom" | "Stretch";

export interface CallbackInfo {
    callbackEndpoint: string;
    callbackInstance?: string;
    callbackApp?: string;
    callbackId?: string;
    targetType?: "All" | "Any";
}

export interface SubscriptionInfo extends CallbackInfo {
    lifetime?: LifetimeType;
}

export type DataOrientation = "Horizontal" | "Vertical";

export interface RangeInfo {
    workbook?: string;
    worksheet?: string;
    range?: string;

    numberFormat?: string;
    expandRange?: boolean;
    resizeOrientation?: DataOrientation;

    /** Ensure the workbook and worksheet exist. If they don't, they will be created. */
    forceCreate?: boolean;
}

export enum ColumnType {
    Integer = "Integer",
    Text = "Text",
    Boolean = "Boolean",
    DateTime = "DateTime",
    Float = "Float",
    Decimal = "Decimal",
}

export interface DataSource {
    file: string;
    name: string;
    table?: string;
    description?: string;
    columns: Column[];
    primaryKey?: string[];
    data?: Array<object[]>;
    transient?: boolean;
}

export interface Column {
    pk: boolean;
    autoIncrement: boolean;
    name: string;
    type: ColumnType;
    nullable?: boolean;
    defaultValue?: any;
}

export enum XLSaveConflictResolution {
    UserResolution = "xlUserResolution",
    LocalSessionChanges = "xlLocalSessionChanges",
    OtherSessionChanges = "xlOtherSessionChanges"
}

export type XLCallback = (origin: any, ...props: any[]) => void;
export type MenuArgs = { returned: { menu?: any; menuId?: string } };
export type ArgsType = { returned: any };
export type TableArgs = { returned: { subscriptionId?: string } };

export class IOConnectDBService {
    private io: any;
    private methodNs: string;

    constructor(ioInstance: any, methodNs: string = "T42.DB.") {
        this.io = ioInstance;
        this.methodNs = methodNs;
    }

    init(dataSource: DataSource): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}Init`, { dataSource })
            .then((args: any) => args.returned);
    }

    createTable(dataSource: DataSource): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}CreateTable`, { dataSource })
            .then((args: any) => args.returned);
    }

    insertData(dataSource: DataSource): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}InsertData`, { dataSource })
            .then((args: any) => args.returned);
    }

    updateRow(dataSource: DataSource, rowData: object[], pkValue: any): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}UpdateRow`, { dataSource, rowData, pkValue })
            .then((args: any) => args.returned);
    }

    updateColumns(dataSource: DataSource, updates: Record<string, any>, pkValue: any): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}UpdateColumns`, { dataSource, updates, pkValue })
            .then((args: any) => args.returned);
    }

    executeQuery(dataSource: DataSource, query: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}ExecuteQuery`, { dataSource, query })
            .then((args: any) => args.returned);
    }

    dispose(dataSource: DataSource): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}Dispose`, { dataSource })
            .then((args: any) => args.returned);
    }
}

export class IOConnectXLService {
    private io: any;
    private methodNs: string;
    private callbackMap: Map<string, XLCallback>;

    private readonly xlServiceCallback = "xlServiceCxtMenuCallback";

    constructor(ioInstance: any, methodNs: string = "IO.XL.") {
        this.io = ioInstance;
        this.methodNs = methodNs;
        this.callbackMap = new Map();

        this.io.interop.register(this.xlServiceCallback, (args: any) => {
            const { subscriptionId } = args;
            const callback = this.callbackMap.get(subscriptionId);
            if (callback) {
                callback(args);
            } else {
                // TODO: Choose a proper warning mechanism.
                console.warn("Missing callback.")
            }
        });
    }

    createWorkbook(workbookFile: string, worksheet: string, saveConflictResolution: XLSaveConflictResolution = XLSaveConflictResolution.UserResolution): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateWorkbook`, { workbookFile, worksheet, saveConflictResolution })
            .then((args: ArgsType) => args.returned);
    }

    subscribeDeltasRaw(range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}SubscribeDeltas`, { range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    subscribeDeltas(rangeInfo: RangeInfo, callback: XLCallback): Promise<XLServiceResult> {
        return this.subscribeDeltasRaw(rangeInfo, {
            callbackEndpoint: this.xlServiceCallback
        }).then((returned: any) => {
            const subscriptionId = returned.subscriptionId;
            if (subscriptionId) {
                this.callbackMap.set(subscriptionId, callback);
            } else {
                // TODO: Choose a proper warning mechanism.
                console.warn("No subscription ID.")
            }

            return returned;
        });
    }

    subscribeRaw(range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}Subscribe`, { range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    subscribe(rangeInfo: RangeInfo, callback: XLCallback): Promise<XLServiceResult> {
        return this.subscribeRaw(rangeInfo, {
            callbackEndpoint: this.xlServiceCallback
        }).then((returned: any) => {
            const subscriptionId = returned.subscriptionId;
            if (subscriptionId) {
                this.callbackMap.set(subscriptionId, callback);
            } else {
                // TODO: Choose a proper warning mechanism.
                console.warn("No subscription ID.")
            }

            return returned;
        });
    }

    destroySubscription(subscriptionId: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}DestroySubscription`, { subscriptionId })
            .then((args: ArgsType) => args.returned);
    }

    read(range: RangeInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}Read`, { range })
            .then((args: ArgsType) => args.returned);
    }

    write(range: RangeInfo, value: object): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}Write`, { range, value })
            .then((args: ArgsType) => args.returned);
    }

    createTable(range: RangeInfo, tableName: string, tableStyle: string, columns: string[], value: object[][],
        callback: XLCallback): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateTable`, {
            range, tableName, tableStyle, columns, value, subscriptionInfo: {
                callbackEndpoint: this.xlServiceCallback
            }
        })
            .then((args: TableArgs) => {
                const subscriptionId = args.returned.subscriptionId;
                if (subscriptionId) {
                    this.callbackMap.set(subscriptionId, callback);
                }
                return args.returned;
            });
    }

    createLinkedTable(range: RangeInfo, dataSource: DataSource, subscriptionInfo: SubscriptionInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateLinkedTable`, { range, dataSource, subscriptionInfo })
            .then((args: any) => args.returned);
    }

    refreshTable(range: RangeInfo, tableName: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}RefreshTable`, { range, tableName })
            .then((args: ArgsType) => args.returned);
    }

    writeTableRows(range: RangeInfo, tableName: string, rowPosition: number | null, value: object[][]): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}WriteTableRows`, { range, tableName, rowPosition, value })
            .then((args: ArgsType) => args.returned);
    }

    readTableRows(range: RangeInfo, tableName: string, fromRow: number, rowsToRead?: number): Promise<XLServiceResult> {

        //TODO: Default fromRow to 1
        //TODO for stas check rowsToRead against the table size
        return this.io.interop.invoke(`${this.methodNs}ReadTableRows`, { range, tableName, fromRow, rowsToRead })
            .then((args: ArgsType) => args.returned);
    }

    updateTableColumns(range: RangeInfo, tableName: string, columnOps: TableColumnOperationDescriptor[]): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}UpdateTableColumns`, { range, tableName, columnOps })
            .then((args: ArgsType) => args.returned);
    }

    describeTableColumns(range: RangeInfo, tableName: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}DescribeTableColumns`, { range, tableName })
            .then((args: ArgsType) => args.returned);
    }

    readRef(reference: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}ReadXlRef`, { reference })
            .then((args: ArgsType) => args.returned);
    }

    writeRef(reference: string, value: object): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}WriteXlRef`, { reference, value })
            .then((args: ArgsType) => args.returned);
    }

    saveAs(range: RangeInfo, fileName: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}SaveAs`, { range, fileName })
            .then((args: ArgsType) => args.returned);
    }

    openWorkbook(fileName: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}OpenWorkbook`, { fileName })
            .then((args: ArgsType) => args.returned);
    }

    createContextMenuRaw(caption: string, menuPath: string[], range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateContextMenu`, { caption, menuPath, range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    destroyContextMenuRaw(menuId: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}DestroyContextMenu`, { menuId })
            .then((args: ArgsType) => args.returned);
    }

    writeComment(range: RangeInfo, comment: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}WriteComment`, { range, comment })
            .then((args: ArgsType) => args.returned);
    }

    clearComments(range: RangeInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}ClearComments`, { range })
            .then((args: ArgsType) => args.returned);
    }

    clearContents(range: RangeInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}ClearContents`, { range })
            .then((args: ArgsType) => args.returned);
    }

    applyStyles(range: RangeInfo, backgroundColor: string, foregroundColor: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}ApplyStyles`, { range, backgroundColor, foregroundColor })
            .then((args: ArgsType) => args.returned);
    }

    setRangeFormat(range: RangeInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}SetRangeFormat`, { range })
            .then((args: ArgsType) => args.returned);
    }

    createDynamicRibbonMenuRaw(caption: string, range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateDynamicRibbonMenu`, { caption, range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    createOrUpdateCTPRaw(
        range: RangeInfo,
        ctpDescriptor: CTPDescriptor,
        subscriptionInfo: SubscriptionInfo
    ): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateOrUpdateCTP`, {
            range,
            ctpDescriptor,
            subscriptionInfo
        }).then((args: ArgsType) => args.returned.result);
    }

    createOrUpdateCTP(
        range: RangeInfo,
        ctpDescriptor: CTPDescriptor,
        callback: XLCallback
    ): Promise<XLServiceResult> {
        const subscriptionInfo: SubscriptionInfo = {
            callbackEndpoint: this.xlServiceCallback,
            callbackId: ctpDescriptor.id
        };

        const overrideCallbackEndpoint = (ui: UIDescriptor) => {
            if (ui.type === "Button" && !ui.callback?.callbackEndpoint) {
                ui.callback = {
                    ...ui.callback,
                    callbackEndpoint: this.xlServiceCallback
                };
            }
            ui.children?.forEach(overrideCallbackEndpoint);
        };

        overrideCallbackEndpoint(ctpDescriptor.ui);

        return this.createOrUpdateCTPRaw(range, ctpDescriptor, subscriptionInfo)
            .then(result => {
                const id = ctpDescriptor.id;
                if (id) {
                    this.callbackMap.set(id, callback);
                }
                return result;
            });
    }

    createDynamicRibbonMenu(
        caption: string,
        range: RangeInfo,
        callback: XLCallback,
    ): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateDynamicRibbonMenu`, {
            caption, range, subscriptionInfo: {
                callbackEndpoint: this.xlServiceCallback
            }
        })
            .then((args: MenuArgs) => {
                const menuId = args.returned.menuId;
                if (menuId) {
                    this.callbackMap.set(menuId, callback);
                }
                return args.returned;
            });
    }

    destroyRibbonMenuRaw(menuId: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}DestroyRibbonMenu`, { menuId })
            .then((args: ArgsType) => args.returned);
    }

    destroyRibbonMenu(menuId: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}DestroyRibbonMenu`, { menuId })
            .then((args: MenuArgs) => {
                // Remove the callback from the map when menu is destroyed
                this.callbackMap.delete(menuId);
                return args.returned;
            });
    }

    createContextMenu(
        caption: string,
        menuPath: string[],
        range: RangeInfo,
        callback: XLCallback,
    ): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}CreateContextMenu`, {
            caption, menuPath, range, subscriptionInfo: {
                callbackEndpoint: this.xlServiceCallback
            }
        })
            .then((args: MenuArgs) => {
                const menuId = args.returned.menuId;
                if (menuId) {
                    this.callbackMap.set(menuId, callback);
                }
                return args.returned;
            });
    }

    destroyContextMenu(menuId: string): Promise<XLServiceResult> {
        return this.io.interop.invoke(`${this.methodNs}DestroyContextMenu`, { menuId })
            .then((args: MenuArgs) => {
                // Remove the callback from the map when menu is destroyed
                this.callbackMap.delete(menuId);
                return args.returned;
            });
    }

    activate(range?: RangeInfo): Promise<XLServiceResult> {
        return this.io.interop
            .invoke(`${this.methodNs}Activate`, { range })
            .then((args: { returned: { result?: any } }) => args.returned);
    }
}