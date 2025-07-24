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
    ],
};
await glueDBService.init(dataSource);
xlService.createLinkedTable({workbook: 'Book3', worksheet: 'Sheet1', range: 'C16', forceCreate: true}, dataSource, undefined).then(console.info)

// loading
glueDBService.init({name: "UserTable",
    file: 'interop.io/io.Connect Desktop/UserData/DEMO-INTEROP.IO/io.db'}).then(console.info)

// inserting 
dataSource.data = [[null, "Mol Gad", "molgad@example.com"]]
glueDBService.insertData(dataSource).then(console.info)

// query
await glueDBService.executeQuery(dataSource, 'select * from UserTable')

// create context menu for any sheet, which works for A1:B5
xlService.createContextMenu("send", ["io","actions"], {range: "A1:B5"}, console.info)

// create dynamic ribbon menu - add in a dropdown that can be executed from the ribbon
xlService.createDynamicRibbonMenu('Another', {range: 'A1:B10'}, console.info)
*/

interface TableColumnOp {
    OldName: string;
    Name: string;
    Position: number | null;
    Op: 'Add' | 'Delete' | 'Rename' | 'Update';
}

enum LifetimeType {
    None = "None",
    GlueInstance = "GlueInstance",
    Forever = "Forever",
    ExcelSession = "ExcelSession"
}

interface SubscriptionInfo {
    callbackEndpoint?: string;
    callbackInstance?: string;
    callbackApp?: string;
    targetType?: string; // Optional, default handled externally
    callbackId?: string;
    lifetime?: LifetimeType; // Optional, default handled externally
}

interface RangeInfo {
    workbook: string;
    worksheet: string;
    range: string;
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


type XlCallback = (origin: any, ...props: any[]) => void;
type MenuArgs = { returned: { menu?: any; menuId?: string } };
type ArgsType = { returned: any };
type TableArgs = { returned: { subscriptionId?: string } };

export class GlueDBService {
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

export class GlueExcelService {
    private io: any;
    private methodNs: string;
    private callbackMap: Map<string, XlCallback>;

    private readonly xlServiceCallback = 'xlServiceCxtMenuCallback';

    constructor(ioInstance: any, methodNs: string = 'T42.XL.') {
        this.io = ioInstance;
        this.methodNs = methodNs;
        this.callbackMap = new Map();

        this.io.interop.register(this.xlServiceCallback, (args: any) => {
            const { origin, subscriptionId, ...otherProps } = args;
            const callback = this.callbackMap.get(subscriptionId);
            if (callback) {
                callback(origin, subscriptionId, ...Object.values(otherProps));
            }
        });
    }

    createWorkbook(workbookFile: string, worksheet: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}CreateWorkbook`, { workbookFile, worksheet })
            .then((args: ArgsType) => args.returned);
    }

    subscribeDeltas(range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}SubscribeDeltas`, { range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    subscribe(range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}Subscribe`, { range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    destroySubscription(subscriptionId: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}DestroySubscription`, { subscriptionId })
            .then((args: ArgsType) => args.returned);
    }

    read(range: RangeInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}Read`, { range })
            .then((args: ArgsType) => args.returned);
    }

    write(range: RangeInfo, value: object): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}Write`, { range, value })
            .then((args: ArgsType) => args.returned);
    }

    createTable(range: RangeInfo, tableName: string, tableStyle: string, columns: string[], value: object[][],
        callback: XlCallback): Promise<object> {
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

    createLinkedTable(range: RangeInfo, dataSource: DataSource, subscriptionInfo: SubscriptionInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}CreateLinkedTable`, { range, dataSource, subscriptionInfo })
            .then((args: any) => args.returned);
    }

    refreshTable(range: RangeInfo, tableName: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}RefreshTable`, { range, tableName })
            .then((args: ArgsType) => args.returned);
    }

    writeTableRows(range: RangeInfo, tableName: string, rowPosition: number | null, value: object[][]): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}WriteTableRows`, { range, tableName, rowPosition, value })
            .then((args: ArgsType) => args.returned);
    }

    readTableRows(range: RangeInfo, tableName: string, fromRow: number, rowsToRead?: number): Promise<object> {

        //TODO: Default fromRow to 1
        //TODO for stas check rowsToRead against the table size
        return this.io.interop.invoke(`${this.methodNs}ReadTableRows`, { range, tableName, fromRow, rowsToRead })
            .then((args: ArgsType) => args.returned);
    }

    updateTableColumns(range: RangeInfo, tableName: string, columnOps: TableColumnOp[]): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}UpdateTableColumns`, { range, tableName, columnOps })
            .then((args: ArgsType) => args.returned);
    }

    describeTableColumns(range: RangeInfo, tableName: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}DescribeTableColumns`, { range, tableName })
            .then((args: ArgsType) => args.returned);
    }

    readXlRef(reference: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}ReadXlRef`, { reference })
            .then((args: ArgsType) => args.returned);
    }

    writeXlRef(reference: string, value: object): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}WriteXlRef`, { reference, value })
            .then((args: ArgsType) => args.returned);
    }

    saveAs(range: RangeInfo, fileName: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}SaveAs`, { range, fileName })
            .then((args: ArgsType) => args.returned);
    }

    openWorkbook(fileName: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}OpenWorkbook`, { fileName })
            .then((args: ArgsType) => args.returned);
    }

    createContextMenuRaw(caption: string, menuPath: string[], range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}CreateContextMenu`, { caption, menuPath, range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    destroyContextMenuRaw(menuId: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}DestroyContextMenu`, { menuId })
            .then((args: ArgsType) => args.returned);
    }

    writeComment(range: RangeInfo, comment: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}WriteComment`, { range, comment })
            .then((args: ArgsType) => args.returned);
    }

    clearComments(range: RangeInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}ClearComments`, { range })
            .then((args: ArgsType) => args.returned);
    }

    clearContents(range: RangeInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}ClearContents`, { range })
            .then((args: ArgsType) => args.returned);
    }

    applyStyles(range: RangeInfo, backgroundColor: string, foregroundColor: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}ApplyStyles`, { range, backgroundColor, foregroundColor })
            .then((args: ArgsType) => args.returned);
    }

    createDynamicRibbonMenuRaw(caption: string, range: RangeInfo, subscriptionInfo: SubscriptionInfo): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}CreateDynamicRibbonMenu`, { caption, range, subscriptionInfo })
            .then((args: ArgsType) => args.returned);
    }

    createDynamicRibbonMenu(
        caption: string,
        range: RangeInfo,
        callback: XlCallback,
    ): Promise<object> {
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
                return args.returned.menu;
            });
    }

    destroyRibbonMenuRaw(menuId: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}DestroyRibbonMenu`, { menuId })
            .then((args: ArgsType) => args.returned);
    }

    destroyRibbonMenu(menuId: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}DestroyRibbonMenu`, { menuId })
            .then((args: MenuArgs) => {
                // Remove the callback from the map when menu is destroyed
                this.callbackMap.delete(menuId);
                return args.returned.menu;
            });
    }

    createContextMenu(
        caption: string,
        menuPath: string[],
        range: RangeInfo,
        callback: XlCallback,
    ): Promise<object> {
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
                return args.returned.menu;
            });
    }

    destroyContextMenu(menuId: string): Promise<object> {
        return this.io.interop.invoke(`${this.methodNs}DestroyContextMenu`, { menuId })
            .then((args: MenuArgs) => {
                // Remove the callback from the map when menu is destroyed
                this.callbackMap.delete(menuId);
                return args.returned.menu;
            });
    }

    activate(range?: RangeInfo): Promise<object> {
        return this.io.interop
            .invoke(`${this.methodNs}Activate`, { range })
            .then((args: { returned: { result?: any } }) => args.returned);
    }
}