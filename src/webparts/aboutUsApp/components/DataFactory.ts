import { WebPartContext } from "@microsoft/sp-webpart-base";
import { find, trim } from "lodash";

// https://pnp.github.io/pnpjs
//> npm install --save @pnp/sp @pnp/graph
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/forms";
import { IList, IListEnsureResult, IListInfo } from "@pnp/sp/lists";
import { IFieldAddResult, IFieldInfo, IFieldUpdateResult } from "@pnp/sp/fields";
import { IViewAddResult, IViewInfo, IViews, IViewUpdateResult } from "@pnp/sp/views";
import { IForms } from "@pnp/sp/forms";

export interface IDataFactoryFieldInfo extends IFieldInfo {
    MaxLength?: number;
    Choices?: string[];
    NumberOfLines?: number;
    RichText?: boolean;
    MaximumValue?: number;
    MinimumValue?: number;
    LookupField?: string;
    LookupList?: string;
    LookupWebId?: string;
    DependentLookupInternalNames?: string[];
    DateTimeCalendarType?: number;
    DisplayFormat?: number;
    FriendlyDisplayFormat?: number;
    AllowMultipleValues?: boolean;
    SelectionGroup?: number;
    SelectionMode?: number;
}

export interface IFieldUrlValue {
    __metadata: {type: "SP.FieldUrlValue"};
    Description: string;
    Url: string;
}

export interface IUserInfoItem {
    __metadata: {type: "SP.Data.UserInfoItem", id?: string};
    ID: number;
    Name?: string;
    Title?: string;
}

export interface IAboutUsListViewTemplate {
    "Title": string;
    "PersonalView": boolean;
    "settings": Partial<IViewInfo>;
    "ViewFields": string[];      // Array of field InternalNames
}
export interface IAboutUsListTemplate {
    "list": {
        "Description": string,
        "Template": number,
        "EnableContentTypes": boolean,
        "settings": Partial<IListInfo>
    };
    "fields": Partial<IListInfo>[] | any;
    "views": IAboutUsListViewTemplate[];
}
export interface IListValidationResults {
    valid: boolean;
    message: string;
}
interface IFieldStatusResults {
    "exists": boolean;
    "update": Partial<IDataFactoryFieldInfo>;
}

//> npm install --save debug
//> npm install --save-dev @types/debug

// private list = new DataFactory("AboutUs_Content");
// await list.ensure(); // ensures list exists. create list/fields if missing
/**
 * About Us Data CRUD operations.
 * @example
 * 
 */
export default class DataFactory {
    //#region PROPERTIES
    private ctx: WebPartContext;
    private requestDigest_: {"value": string, "expires": Date} = {
        "value": null,
        "expires": null,
    };

    // List Title - listName
    private title_: string = "";
    public get title() : string {
        return this.title_;
    }
    public set title(v : string) {
        this.title_ = v;
        this.api_ = sp.web.lists.getByTitle(v);
        this.restApi_ = `${ this.ctx.pageContext.web.absoluteUrl }/_api/web/lists/getByTitle('${ v }')`;
    }

    /* How to get list 'fields' template from an existing SP List.
        fetch(_spPageContextInfo.webServerRelativeUrl + "_api/web/lists/getByTitle('[LIST_NAME]')/fields?$filter=(CanBeDeleted eq true) or (InternalName eq 'Title')",
            { headers: new Headers({ accept: "application/json;odata=verbose;" }) })
            .then( response => response.json() )
            .then( json => console.info(json.d.results) );
    */
    // list template
    private listTemplate: IAboutUsListTemplate = require('./AboutUsListTemplate.json');
    
    // api
    private restApi_: string;
    public get restApi(): string {
        return this.restApi_;
    }

    private api_: IList;
    public get api(): IList {
        return this.api_;
    }

    // list info
    private properties_: IListInfo;
    public get properties(): IListInfo {
        return this.properties_;
    }

    // fields
    private fields_: IDataFactoryFieldInfo[] = [];
    public get fields() : IDataFactoryFieldInfo[] {
        return this.fields_;
    }

    // views
    private views_: IViews = null;
    public get views() : IViews {
        return this.views_;
    }

    // forms
    private forms_: IForms = null;
    public get forms() : IForms {
        return this.forms_;
    }
    
        
    // exists
    private exists_: boolean = false;
    public get exists() : boolean {
        return this.exists_;
    }
    
    //#endregion

    //#region CONSTRUCTOR
    constructor(context: WebPartContext) {
        // we need the current webpart context to make REST requests against. 
        this.ctx = context;
        sp.setup({ spfxContext: context });
    }
    //#endregion

    //#region LIST
    /**
     * Ensure list and fields exists. Creates list and fields if missing.
     * Call this to set the 'title' property.
     * @param title List title to ensure exists.
     */
    public async ensureList(title: string): Promise<void> {
        title =  trim(title);

        // make sure list title is valid
        if (!DataFactory.validateListName(title).valid) {
            this.exists_ = false;
            return;
        }

        // create list if it doesn't exist.
        // SP-PnP check will throw a '404' status if the list doesn't exists.
        const ensureResults: IListEnsureResult = await sp.web.lists.ensure(
                title,
                `This list was auto-generated by the 'About Us' app on ${ (new Date()).toLocaleString() }. \r\n ${ this.listTemplate.list.Description }`,
                this.listTemplate.list.Template,
                this.listTemplate.list.EnableContentTypes,
                this.listTemplate.list.settings
            );
        await this.sleep(50);   // pause to let SPO do the work
        
        // set title (initializes the this.api and this.restApi)
        this.title = title;

        // ensure list fields
        await this.ensureFields();

        // ensure views after the list and fields have been created/updated
        await this.ensureViews();

        // get form: display, edit, new
        this.forms_ = await this.getAllForms();

        // get list properties
        this.exists_ = true;
        this.properties_ = await this.api.expand("DefaultView", "Forms", "RoleAssignments", "RootFolder").get({ "headers": { "Accept": "application/json;odata=verbose"} });
    }
    //#endregion

    //#region FIELDS
    /**
     * Check if fields exists. Tries to update or create missing fields.
     */
    public async ensureFields(): Promise<IDataFactoryFieldInfo[]> {

        // list name must be set
        if (this.title === null || this.title === "") return;

        // get existing fields from list
        const existingFields: IDataFactoryFieldInfo[] = await this.getAllFields();

        // loop thru each field. don't use array.forEach, it doesn't honor 'await'.
        for (let i = 0; i < this.listTemplate.fields.length; i++) {
            const field = this.listTemplate.fields[i],
                status = DataFactory.compareTemplateToListField(field, existingFields);

            if (!status.exists) {
                // field doesn't exist, create the field
                await this.createField(field);

            } else if (status.exists && status.update !== null) {
                // field exists, but needs to be updated
                await this.updateField(status.update);

            }

        }

        // get new set of fields.
        this.fields_ = await this.getAllFields();
        return this.fields;
    }

    /**
     * Request all list fields.  Filters out non-deletable fields.  Includes 'Title' field
     * @returns SP PnP array of field information
     */
    private async getAllFields(): Promise<IDataFactoryFieldInfo[]> {
        return this.api.fields
            .filter("(CanBeDeleted eq true) or (InternalName eq 'Title')")
            .get();
    }

    /**
     * Compares field template to existing fields in the list.
     * Get the field status (does the field already exist?, what properties need to be updated).
     * @param field Template Field Object. This is what the list field should look like.
     * @param existingFields Array of existing Field Objects from the list.  This is what the list field looks like now.
     * @returns status.exists = if field already exists; status.update = if any field properties needs to be updated.  
     */
    private static compareTemplateToListField(field: { [prop: string]: any }, existingFields: { [prop: string]: any }[]): IFieldStatusResults {

        // set default status: does not exist & no update
        let status = { "exists": false, "update": null };

        const existingField = find(existingFields, ["InternalName", field.InternalName]);
        // const propertyNameTranslation = { "AutoIndexed": "Indexed" },
        //     translatePropertyName = propertyName => {
        //         return (Object.prototype.hasOwnProperty.call(propertyNameTranslation, propertyName))
        //             ? propertyNameTranslation[propertyName]
        //             : propertyName ;
        //         };
        
        // if there is an existing field, check the field's properties
        if (existingField) {
            status.exists = true;

            const //props = Object.keys(field),
                propertiesToIgnore = ["__metadata", "LookupList", "Description"];

            // loop through the fields properties to find any changes
            Object.keys(field).forEach( prop => {
                let fieldValue = field[prop],
                    existingValue = existingField[prop];
                
                if (propertiesToIgnore.indexOf(prop) > -1) return true; // skip this prop

                // normalize values.  null and "" are the same things
                if (fieldValue === null) fieldValue = "";
                if (existingValue === null) existingValue = "";

                // "Choices" property - only compare dropdown choices (Choices.results array)
                if (prop === "Choices") {
                    fieldValue = JSON.stringify(fieldValue.results);
                    existingValue = JSON.stringify(existingValue);
                }

                if (fieldValue !== existingValue) {
                    // initialize update object?
                    if (status.update === null) status.update = { "Id": existingField.Id };
                    // this property needs to be updated
                    status.update[prop] = field[prop];
                }
            // }
            });
        }

        return status;
    }

    /**
     * Create SP field
     * @param field Field properties. Similar JSON as SP REST response.
     * @returns Promise
     */
    public async createField(field: any): Promise<Partial<IDataFactoryFieldInfo>> {
        let addResult: IFieldAddResult;
        let newFieldInfo: Partial<IDataFactoryFieldInfo> = null;

        // temporarily set the field's title the same as the InternalName. We will have to update the title afterwards.
        const fieldTitle = field.Title || field.InternalName;
        field.Title = field.InternalName;

        // create the field
        try{
            // lookups are handled differently.  we need to get the list and web GUID first
            switch (field.__metadata.type) {
                case "SP.FieldLookup":
                    // keyword "[THISLIST]" refers to the current "AboutUs Contents" list
                    const list: IListInfo = await DataFactory.getList((field.LookupList === "[THISLIST]") ? this.title : field.LookupList );
                    
                    // exit now if list doesn't exist.
                    if (list === null) {
                        // unable to set lookup list because the list doesn't exist
                        DataFactory.DEBUG(`ERROR: createField(): Unable to create lookup field because the list '${ field.LookupList }' doesn't exist!`, field);
                        return null;
                    }

                    // update LookupList
                    field.LookupList = list.Id;

                    // create lookup field
                    addResult = await this.api.fields.addLookup(
                        field.Title,
                        list.Id,
                        field.LookupField
                    );

                    //  need to update the rest of the lookup field's properties
                    await this.updateField({
                        "Id": addResult.data.Id,
                        "Description": field.Description || "",
                        "Hidden": field.Hidden || false,
                        "Required": field.Required || false
                    });

                    break;
            
                default:
                    // create list field/column
                    addResult = await this.api.fields.add(
                        field.Title,
                        field.__metadata.type,
                        field
                    );
                
                    break;
            }

        } catch(er) {
            DataFactory.DEBUG("ERROR: createField():", field, er);
            return null;
        }

        newFieldInfo = addResult.data;
        await this.sleep(50); // pause to let SPO do the work

        // update Display Title if different than internal name
        if (fieldTitle !== newFieldInfo.Title) {
            const updateFieldInfo = {
                    "__metadata": {
                        "type": field.__metadata.type
                    },
                    "Id": newFieldInfo.Id,
                    "Title": fieldTitle
                };
            
            const updateResult: IFieldUpdateResult = await this.updateField(updateFieldInfo);
            newFieldInfo = updateResult.data;

        }

        // add field to default view?

        return newFieldInfo;
    }

    /**
     * Update's SP field. Requires field Id (GUID)
     * @param field Field properties. Similar JSON as SP REST response. Must include: __metadata.type & Id
     */
    public async updateField(field: Partial<IDataFactoryFieldInfo>): Promise<IFieldUpdateResult> {
        let updateResult: IFieldUpdateResult = null;

        try {
            updateResult = await this.api.fields.getById(field.Id).update(field);

        } catch(er) {
            DataFactory.DEBUG("ERROR: updateField():", field, er);
        }

        return updateResult;
    }

    /**
     * Get the list of colors choices from the Org-Chart Colors.
     * @returns Array of color names.
     */
    public getOrgChartColors(): string[] {
        let colors = [];
        const exists = this.exists;

        if (exists && this.fields.length > 0) {
            const orgChartColorField = <any>find(this.fields, ["InternalName", "OrgChartColor"]);
            colors = orgChartColorField.Choices.map( color => color.split(/ ?\(/)[0] );
        }

        return colors;
    }
    //#endregion

    //#region VIEWS
    /**
     * Updates existing views with changes or if the view doesn't exist, creates new view.
     */
    private async ensureViews(): Promise<void> {
        const existingViews: IViews = await this.getAllViews();

        // loop through each view in template
        this.listTemplate.views.forEach( async viewTemplate => {
            const existingView = <IViewInfo>find(existingViews, ['Title', viewTemplate.Title]);

            if (existingView === undefined || existingView === null) {
                // view not found, create view.
                await this.createView(viewTemplate);
                
            } else {
                // view exists. check properties
                let viewUpdateInfo: Partial<IViewInfo> = null;

                // loop through each view template properties
                Object.keys(viewTemplate.settings).forEach( prop => {
                    const templateValue = viewTemplate.settings[prop],
                        viewValue = existingView[prop];

                    if ( templateValue !== viewValue ) {
                        if  (viewUpdateInfo === null) viewUpdateInfo = {};
                        viewUpdateInfo[prop] = templateValue;
                    }
                });

                // update view
                if (viewUpdateInfo !== null) {
                    await this.updateView(existingView.Id, viewUpdateInfo, viewTemplate.Title);
                }

                await this.ensureViewFields(existingView.Id, viewTemplate.ViewFields);
            }
        });

        this.views_ = await this.getAllViews();
    }

    /**
     * Get all current list views.
     * @returns Array of list views
     */
    private async getAllViews(): Promise<IViews> {
        return await this.api.views();
    }

    /**
     * Create a new list view.
     * @param viewInfo List view settings/properties.
     * @returns Results of creation request
     */
    private async createView(viewInfo: Partial<IAboutUsListViewTemplate>): Promise<IViewAddResult> {
        let addResult: IViewAddResult = null;
        
        try {
            addResult = await this.api.views.add(
                viewInfo.Title,
                viewInfo.PersonalView || false,
                viewInfo.settings || {}
            );

            await this.ensureViewFields(addResult.data.Id, viewInfo.ViewFields);
            await this.sleep(50);   // pause to let SPO do the work
                        
        } catch(er) {
            DataFactory.DEBUG(`ERROR! Unable to update '${ viewInfo.Title }' view.`, viewInfo, er);
        }

        return addResult;
    }

    /**
     * Update a list view properties. Use ensureViewFields() to update view fields.
     * @param viewId GUID for the list view to update.
     * @param viewInfo View settings/properties to update.
     * @param viewTitle Title of the list view to update. Used for debug messages only.
     * @returns Results of the update request
     */
    private async updateView(viewId: string, viewInfo: Partial<IViewInfo>, viewTitle: string = ""): Promise<IViewUpdateResult> {
        let updateResult: IViewUpdateResult = null;
        
        try {
            updateResult = await this.api.views.getById(viewId).update(viewInfo);

        } catch(er) {
            DataFactory.DEBUG(`ERROR! Unable to update '${ viewTitle || viewId }' view.`, viewInfo, er);
        }

        return updateResult;
    }

    /**
     * Ensure only these fields are in the view and in a specific order.  Removes all other fields.
     * @param viewId GUID for the list view to update.
     * @param fieldNames Array of internal or display field names. Order of fields is taken into account.
     */
    private async ensureViewFields(viewId: string, fieldNames: string[]): Promise<void> {
        const fields = this.api.views.getById(viewId).fields,
            viewFields = await <Promise<{[key:string]: any}>>fields.get(),
            existingFieldNames = viewFields.Items;

        // if existing field names and order match template, there's nothing to update.
        if (JSON.stringify(existingFieldNames) === JSON.stringify(fieldNames)) return;

        // remove all fields
        try {
            await fields.removeAll();   // NOTE: SPO resolves the request eventhough the backend isn't done
            await this.sleep(1000);  // pause (for awhile!) to let SPO do the work


        } catch(er) {
            DataFactory.DEBUG('ERROR! Unable to remove fields from view.', `View ID: ${ viewId }`, er);
        }

        // add fields in proper order. don't use array.forEach, it doesn't honor 'await'.
        for (let i = 0; i < fieldNames.length; i++) {
            const fieldName = fieldNames[i];
            
            try {
                await fields.add(fieldName);   // NOTE: SPO resolves the request eventhough the backend isn't done
                await this.sleep(100);  // pause to let SPO do the work. If you don't pause here, the view field may not get added

            } catch(er) {
                DataFactory.DEBUG(`ERROR! Unable to add '${ fieldName }' field to view.`, `View ID: ${ viewId }`, er);
            }
        }
    }
    //#endregion

    //#region FORMS
    /**
     * Get all the list forms.
     * @returns List of IFormInfo for Display, Edit & New forms.
     */
    private async getAllForms(): Promise<IForms> {
        return this.api.forms.get();
    }
    //#endregion

    //#region ITEMS
    //#endregion

    //#region HELPERS
    /**
     * Get list properties by list title.
     * @param title Title for list to get
     * @returns List properties. Returns null if not found.
     */
     private static async getList(title: string, ...select: string[]): Promise<IListInfo> {
        try {
            // will throw error if list doesn't exist
            const list = sp.web.lists.getByTitle(title);

            return await list
                .select.apply(list, select)
                .get({ "headers": { "Accept": "application/json;odata=nometadata" } });
        
        } catch(er) {
            return null;

        }
    }

    /**
     * Checks to see if list exists. Checks by list title.
     * Same request as PnP list.ensure(), except this will not automatically create the list.
     * @param title List title to check.
     * @returns True if list exists.
     */
    public static async listExists(title: string): Promise<boolean> {
        title = trim(title);

        if (!DataFactory.validateListName(title).valid) return false;

        // valid list name, check to see if list exists. will throw '404' if list doesn't exist.
        // same request as PnP ensure, except this doesn't automatically create the list.
        const list = await DataFactory.getList(title, "Title");
        return list !== null;
    }
    
    /**
     * Get the SP Web's RequestDigest.
     * @param renew Requests the RequestDigest again, ignores cache
     * @returns Request digest value
     */
    public async getRequestDigest(renew: boolean = false): Promise<string> {
        const ePageDigest = (<HTMLInputElement>document.getElementById('__REQUESTDIGEST')),
            pageDigest = (ePageDigest) ? ePageDigest.value : null,
            NOW = new Date(),
            isExpired = (this.requestDigest_.expires === null) ? true : NOW >= this.requestDigest_.expires;

        if (pageDigest !== null) {
            // use existing request digest
            this.requestDigest_.value = pageDigest;

        } else if (this.requestDigest_.value === null || renew === true || isExpired) {
            // get new request digest using the REST API
            try{
                const response = await fetch(this.ctx.pageContext.web.absoluteUrl + "/_api/contextinfo", {
                        "method": "POST",
                        "headers": new Headers({ "Accept": "application/json;odata=verbose" })
                    }),
                    json = await response.json(),
                    data = json.d.GetContextWebInformation,
                    value = data.FormDigestValue,
                    timeoutSeconds = data.FormDigestTimeoutSeconds,
                    date = new Date(value.split(",")[1]);

                // update expiration time
                date.setTime(date.getTime() + timeoutSeconds * 1000);
                this.requestDigest_.expires = NOW;
                this.requestDigest_.value = data.FormDigestValue;

            } catch(er) {
                this.requestDigest_.expires = null;
                this.requestDigest_.value = null;
            }
        }

        return this.requestDigest_.value;
    }

    /**
     * REST request to retrieve all SP lists from site.
     * @param select Array of REST $select values. Default: ["Title"]
     * @param expand Array of REST $expand values. Default: []
     * @param filter REST $filter value. Default: "(BaseTemplate eq 100) and (Hidden eq false)"
     * @returns Array of list objects.
     */
    public static async getAllLists(
        select: string[] = ["Title"], 
        expand: string[] = [],  
        filter: string = "(BaseTemplate eq 100) and (Hidden eq false)"
        ): Promise<any> {

        const LISTS = sp.web.lists;

        const DATA = await LISTS
            .select.apply(LISTS, select)
            .expand.apply(LISTS, expand)
            .filter(filter)
            .get({
                "headers": {"Accept": "application/json;odata=nometadata"}
            });

        return DATA;
    }

    /**
     * Check to see if list name is valid.
     * @param name List name to check.
     * @param existingListNames Array of existing list names. Use DataFactory.GetAllLists() to retrieve all list titles.
     * @returns Object: { "valid": boolean, "message": string }
     */
    public static validateListName(name: string, existingListNames: string[] = []): IListValidationResults {
        name = trim(name);

        // if null or empty = not valid
        if (name === null || name.length === 0) {
            return { "valid": false, "message": "Must provide a list name. 3 characters minimum." };
        }

        // if less than 3 characters
        if (name.length < 3) {
            return { "valid": false, "message": "3 characters minimum." };
        }

        // special characters, except underscores and parathesis'
        if ((/^[a-zA-Z0-9-_() ]+$/gi).test(name) === false) {
            return { "valid": false, "message": "List name cannot contain special characters.\nUnderscores, parathisis' and dashes are allowed." };
        }
        
        // if list exists = not valid;
        if (existingListNames.indexOf(name) > -1) {
            return { "valid": false, "message": "List already exists!\nEnter a new list name\nor select from the list dropdown." };
        }

        return { "valid": true, "message": "" };
    }

    /**
     * Pauses the script for a set amount of time.
     * @param milliseconds Amount of milliseconds to sleep.
     * @returns Promise
     * @example
     * await sleep(1000);  // sleep for 1 second then continue
     * // or
     * sleep(500).then(() => {});  // sleep for half second then run function
     */
    private sleep(milliseconds: number): Promise<void> {
        return new Promise(resolve => setTimeout(resolve, milliseconds));
    }

    /**
     * Prints our debug messages. Decorated console.info() or console.error() method.
     * @param args Message or object to view in the console. If message starts with "ERROR", DEBUG will use console.error().
     */
    public static DEBUG(...args: any[]) {
        // is an error message, if first argument is a string and contains "error" string.
        const isError = (args.length > 0 && (typeof args[0] === "string")) ? args[0].toLowerCase().indexOf("error") > -1 : false;
        args = ["(About-Us DataFactory.ts)"].concat(args);

        if (window && window.console) {
            if (isError && console.error) {
                console.error.apply(null, args);

            } else if (console.info) {
                console.info.apply(null, args);

            }
        }
    }
    //#endregion

    public async test(...args: any[]): Promise<any> {
        if (args.length > 0) DataFactory.DEBUG.apply(null, args);
    }
}