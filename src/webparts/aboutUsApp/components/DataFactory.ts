// About-Us data requests to/from the content list and web
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
import "@pnp/sp/site-groups/web";
import "@pnp/sp/security/list";
import * as spLists from "@pnp/sp/lists";
import { IList } from "@pnp/sp/lists";
import { IFieldAddResult, IFieldInfo, IFieldUpdateResult } from "@pnp/sp/fields";
import { IViewAddResult, IViewInfo, IViews, IViewUpdateResult } from "@pnp/sp/views";
import { IForms } from "@pnp/sp/forms";
import { IItem, IItems } from "@pnp/sp/items";
import { ISiteGroupInfo, ISiteGroups } from "@pnp/sp/site-groups/types";
import { IBasePermissions, IRoleAssignmentInfo, IRoleDefinitionInfo, PermissionKind } from "@pnp/sp/security";
import "@pnp/sp/site-users/web";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import "@pnp/sp/search";
import { ISearchQuery, ISearchResult, Search, SearchResults } from "@pnp/sp/search";
import { IDropdownOption, IIconProps, IStyle } from "office-ui-fabric-react";
import { DEBUG_NOTRACE, escapeProperly, IAboutUsAppWebPartProps, LOG, sleep } from "../AboutUsAppWebPart";
import { IFilePickerProps } from "@pnp/spfx-controls-react";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { Folder, IFolder, IFolderInfo } from "@pnp/sp/folders/types";
import { IFileAddResult } from "@pnp/sp/files";
import { IDocumentLibraryInformation } from "@pnp/sp/sites";

//#region INTERFACES, TYPES & ENUMS
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

export interface IListInfo extends spLists.IListInfo {
    [key: string]: any;
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
export interface IAboutUsListRoleDef {
    name: string;
    description: string;
    order: number;
    basePermissions: IBasePermissions;
}
export interface IAboutUsMicroFormField {
    hidden?: boolean;

    type: string;
    internalName: string;
    defaultValue: any;

    id?: string;
    label?: string;
    required?: boolean;
    disabled?: boolean;
    placeholder?: string;
    styles?: Record<string, string>;
    className?: string;
    description?: string;

    options?: IDropdownOption[];
    target?: "_blank" | "_self" | "_parent" | "_top";
    iconProps?: IIconProps;
    mask?: string;
    readOnly?: boolean;
    multiline?: boolean;
    filePickerProps?: Partial<IFilePickerProps>;
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
    "roleDefs": Record<string, IAboutUsListRoleDef>;
    "fieldMicroForms": Record<string, IAboutUsMicroFormField>[];
}

export interface IListValidationResults {
    valid: boolean;
    message: string;
}
interface IFieldStatusResults {
    "exists": boolean;
    "update": Partial<IDataFactoryFieldInfo>;
}

export type TAboutUsRoleDef = "Full Control" | "Contribute" | "Content Manager" | "Read";

export interface IDataStructureItem {
    ID: number;
    Title: string;
    Name?: string;
    DisplayType: string[];
    OrgType?: string;
    OrderBy: number;
    ParentID?: number | string;
    children: IDataStructureItem[];
    data?: Record<string, any>;
    flags?: Record<string, any>;
}

export interface IUserPermissions {
    canAdd: boolean;
    canEdit: boolean;
    canDelete: boolean;
}

//#endregion


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
        fetch(_spPageContextInfo.webServerRelativeUrl + 
            "_api/web/lists/getByTitle('[LIST_NAME]')/fields?$filter=(CanBeDeleted eq true) or (InternalName eq 'Title')",
            { headers: new Headers({ accept: "application/json;odata=verbose;" }) })
            .then( response => response.json() )
            .then( json => console.info(json.d.results) );
    */
    // list template
    public static listTemplate: IAboutUsListTemplate = require('./AboutUsListTemplate.json');
    
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

    // role assignments
    // private roleAssignments_: IRoleAssignmentInfo[];
    // public get roleAssignments() : IRoleAssignmentInfo[] {
    //     return this.roleAssignments_;
    // }

    // role definition
    private static roleDefinitions: Record<string, IRoleDefinitionInfo> = {};

    
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
     * @param updateList If true, will update/reset list properties as defined by the template
     * @param updateFields If true, will update/reset list fields as defined by the template
     * @param updateViews If true, will update/reset list view(s) as defined by the template
     * @param breakRoleInheritance If true, will break role inheritance
     */
    public async ensureList(title: string, updateList: boolean = false, updateFields: boolean = false, updateViews: boolean = false, breakRoleInheritance: boolean = false): Promise<void> {
        title =  trim(title);

        // make sure list title is valid
        if (!DataFactory.validateListName(title).valid) {
            this.exists_ = false;
            return;
        }

        // create list if it doesn't exist.
        // SP-PnP check will throw a '404' status if the list doesn't exists.
        if (updateList) {
            await sp.web.lists.ensure(
                title,
                `This list was auto-generated by the 'About Us' app on ${ (new Date()).toLocaleString() }. \r\n ${ DataFactory.listTemplate.list.Description }`,
                DataFactory.listTemplate.list.Template,
                DataFactory.listTemplate.list.EnableContentTypes,
                DataFactory.listTemplate.list.settings
            );

            await sleep(100);   // pause to let SPO do the work
            this.exists_ = true;
        } else {
            this.exists_ = await DataFactory.listExists(title);
            
            if (this.exists_ === false) return;
        }
        
        // set title (initializes the this.api and this.restApi)
        this.title = title;

        // break role inheritance (unique permissions)
        if (breakRoleInheritance) await this.api.breakRoleInheritance(true);

        // ensure list fields
        if (updateFields) {
            await this.ensureFields();
        } else {
            this.fields_ = await this.getAllFields();
        }

        // ensure views after the list and fields have been created/updated
        if (updateViews) await this.ensureViews();

        // get list properties
        this.properties_ = await this.api.expand(
            "DefaultView",
            "Forms",
            "RoleAssignments",
            "RoleAssignments/Member",
            "RoleAssignments/RoleDefinitionBindings",
            "RootFolder").get({ "headers": { "Accept": "application/json;odata=verbose"} });
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
        for (let i = 0; i < DataFactory.listTemplate.fields.length; i++) {
            const field = DataFactory.listTemplate.fields[i],
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
    private static compareTemplateToListField(
        field: Record<string, any>,
        existingFields: Record<string, any>[]): IFieldStatusResults {

        // set default status: does not exist & no update
        let status = { "exists": false, "update": null };

        const existingField = find(existingFields, ["InternalName", field.InternalName]);
        
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
                    // keyword "[THISLIST]" refers to the current "AboutUs Contents" list name
                    const title = (field.LookupList === "[THISLIST]") ? this.title : field.LookupList,
                        list = await DataFactory.getListByTitle(title, "Id");
                    
                    // exit now if list doesn't exist.
                    if (list === null) {
                        // unable to set lookup list because the list doesn't exist
                        LOG(`ERROR: createField(): Unable to create lookup field because the list '${ field.LookupList }' \
                            doesn't exist!`, field);
                        return null;
                    }

                    // update LookupList
                    field.LookupList = list.Id;

                    // create lookup field
                    addResult = await this.api.fields.addLookup(
                        field.Title,
                        field.LookupList,
                        field.LookupField
                    );

                    await sleep(50);

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
            LOG("ERROR: createField():", field, er);
            return null;
        }

        newFieldInfo = addResult.data;
        await sleep(50); // pause to let SPO do the work

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
            LOG("ERROR: updateField():", field, er);
        }

        return updateResult;
    }

    public getExistingField_InternalName(internalName: string): IDataFactoryFieldInfo {
        return find(this.fields, ["InternalName", internalName]);
    }
//#endregion

//#region VIEWS
    /**
     * Updates existing views with changes or if the view doesn't exist, creates new view.
     */
    private async ensureViews(): Promise<void> {
        const existingViews: IViews = await this.getAllViews();

        // loop through each view in template
        DataFactory.listTemplate.views.forEach( async viewTemplate => {
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
            await sleep(50);   // pause to let SPO do the work
                        
        } catch(er) {
            LOG(`ERROR! Unable to update '${ viewInfo.Title }' view.`, viewInfo, er);
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
            LOG(`ERROR! Unable to update '${ viewTitle || viewId }' view.`, viewInfo, er);
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
            viewFields = await <Promise<Record<string, any>>>fields.get(),
            existingFieldNames = viewFields.Items;

        // if existing field names and order match template, there's nothing to update.
        if (JSON.stringify(existingFieldNames) === JSON.stringify(fieldNames)) return;

        // remove all fields
        try {
            await fields.removeAll();   // NOTE: SPO resolves the request eventhough the backend isn't done
            await sleep(1000);  // pause (for awhile!) to let SPO do the work


        } catch(er) {
            LOG('ERROR! Unable to remove fields from view.', `View ID: ${ viewId }`, er);
        }

        // add fields in proper order. don't use array.forEach, it doesn't honor 'await'.
        for (let i = 0; i < fieldNames.length; i++) {
            const fieldName = fieldNames[i];
            
            try {
                await fields.add(fieldName);   // NOTE: SPO resolves the request eventhough the backend isn't done
                await sleep(100);  // pause to let SPO do the work. If you don't pause here, the view field may not get added

            } catch(er) {
                LOG(`ERROR! Unable to add '${ fieldName }' field to view.`, `View ID: ${ viewId }`, er);
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
    /** Basic GET request for SP List Item
     * @param id Item ID to retrieve
     * @returns API results
     */
    public async getItemById(id: number): Promise<Record<string, any>> {
        let _item = null;
        try {
            _item = await this.api.items.getById(id).get();

        } catch (er) {
            LOG("ERROR! getItemById(): Unable to fetch item by id:", id, er);
        }
        return _item;
    }

    /** basic GET request for SP List Item with Lookup fields expanded.
     * @param id Item ID to retrieve
     * @returns API results
     */
    public async getItemById_expandFields(id: number): Promise<Record<string, any>> {
        const $select = ["Id", "ID"],
            $expand = [];

        this.fields.forEach(field => {
            switch (field["odata.type"]) {
                case "SP.FieldLookup":
                    $select.push(field.InternalName);
                    $select.push(field.InternalName + "Id");
                    $select.push(field.InternalName + "/ID");
                    $select.push(field.InternalName + "/" + field.LookupField);

                    $expand.push(field.InternalName);
                    break;

                case "SP.FieldUser":
                    $select.push(field.InternalName);
                    $select.push(field.InternalName + "Id");
                    $select.push(field.InternalName + "/ID");
                    $select.push(field.InternalName + "/Title");
                    $select.push(field.InternalName + "/Name");
                    $select.push(field.InternalName + "/EMail");

                    $expand.push(field.InternalName);
                    break;
            
                default:
                    $select.push(field.InternalName);
                    break;
            }
        });

        // if item ID was passed use that first
        let item = null;
        try {
            if (typeof id === "number") {
                const apiItems = this.api.items.getById(id);
                item = await apiItems
                    .select.apply(apiItems, $select)
                    .expand.apply(apiItems, $expand)
                    .get();
            }
        } catch (er) {
            LOG("ERROR! getItemById_expandFields(): Unable to get item by ID with expanded lookup fields.", er);
        }

        return item;
    }

    /**
     * REST request to retrieve all SP lists items.
     * @param select Array of REST $select values. Default: ["ID", "Id", "Title"]
     * @param expand Array of REST $expand values. Default: []
     * @param filter REST $filter value. Default: null
     * @param metadataType oData request type. 
     * @returns Array of list objects.
     */
    public async getAllItems(
        select: string[] = ["ID", "Id", "Title"], 
        expand: string[] = [],
        orderBy: [string, boolean] = ["Id", true],
        metadataType: "verbose" | "nometadata" | "minimal" = "nometadata"
        ): Promise<Record<string, any>[]> {

        const apiItems = this.api.items;

        const items = await apiItems
            .select.apply(apiItems, select)
            .expand.apply(apiItems, expand)
            .orderBy.apply(apiItems, orderBy)
            .top(5000)
            .getAll({
                "headers": {"Accept": `application/json;odata=${metadataType}`}
            });

        return items;
    }
//#endregion

//#region PERMISSIONS
    /** Get Permission Level information by name. Caches roles.
     * @param roleName Permission level names. E.i.: Full Control, Design, Edit, Contribute, Read...
     * @returns Request results
     */
    public static async getWebRoleDefinition(roleName: string): Promise<IRoleDefinitionInfo> {
        try {
            if (!(roleName in DataFactory.roleDefinitions)) {
                DataFactory.roleDefinitions[roleName] = await sp.web.roleDefinitions.getByName(roleName).get();
            }

            return DataFactory.roleDefinitions[roleName];

        } catch (er) {
            LOG(`ERROR! Unable to '${ roleName }' get Web Role Definition.`, er);
        }

        return null;
    }

    /** Create the About-Us role (permission level).
     * @param role "Full Control" | "Contribute" | "Content Manager" | "Read". About-Us role names from AboutUsListTemplate.json.
     * @returns About-Us role definition information - Permission Level
     */
    public async createRole(role: TAboutUsRoleDef): Promise<IRoleDefinitionInfo> {
        try {
            const roleTemplate = DataFactory.listTemplate.roleDefs[role],
                response = await sp.web.roleDefinitions.add(
                    roleTemplate.name,
                    roleTemplate.description,
                    roleTemplate.order,
                    roleTemplate.basePermissions
                );

            // save to cache
            DataFactory.roleDefinitions[roleTemplate.name] = response.data;

            return response.data;

        } catch (er) {
            LOG(`ERROR! Unable to create '${role}' role definition (permission level).`, er);
        }

        return null;
    }

    /** Ensures the About-Us role exists, creates it if missing, then returns the role.
     * @param role "Full Control" | "Contribute" | "Content Manager" | "Read". About-Us role names from AboutUsListTemplate.json.
     * @returns About-Us role definition information - Permission Level
     */
    public async ensureRole(role: TAboutUsRoleDef): Promise<IRoleDefinitionInfo> {
        let roleDef = await DataFactory.getWebRoleDefinition(role);

        if (roleDef === null) roleDef = await this.createRole(role);
        
        return roleDef;
    }

    /** Resets and breaks list permission inheritence. Sets new Full Control and Read permissions.
     * @param ownerId Group/Principle ID being granted Full Control permissions
     * @param readerId Group/Principle ID being granted Read permissions
     */
    public async resetListRoleAssignments(ownerId: number, memberId: number, readerId: number): Promise<void> {
        try {
            // make sure the list inheritance is broken
            await this.api.breakRoleInheritance(false);
            await sleep(50);   // pause to let SPO do the work

            // get existing roles
            const existingRoleAssignments: IRoleAssignmentInfo[] = await this.api.roleAssignments.get();

            // remove all existing roles
            existingRoleAssignments.forEach(async ra => {
                try {
                    const response = await this.api.roleAssignments.getById(ra.PrincipalId).delete();
                    await sleep(10);   // pause to let SPO do the work
                } catch (er) {
                    LOG("ERROR! Unable to delete role assignment for:", ra, er);
                }
            });

            // set owners
            if (typeof ownerId === "number") {
                try {
                    // get role definition
                    const fullControlRoleDef = await this.ensureRole("Full Control");

                    // set group to role
                    await this.api.roleAssignments.add(ownerId, fullControlRoleDef.Id);

                } catch (er) {
                    LOG(`ERROR! Unable to assign 'Full Control' for Group ID: ${ownerId}.`, er);
                }
            }

            // set members (manager)
            if (typeof memberId === "number") {
                try {
                    // get role definition
                    const contributeRoleDef = await this.ensureRole("Contribute");

                    // set group to role
                    await this.api.roleAssignments.add(memberId, contributeRoleDef.Id);

                } catch (er) {
                    LOG(`ERROR! Unable to assign 'Contribute' for Group ID: ${memberId}.`, er);
                }
            }

            // set readers
            if (typeof readerId === "number") {
                try {
                    // get role definition
                    const readRoleDef = await this.ensureRole("Read");

                    // set group to role
                    await this.api.roleAssignments.add(readerId, readRoleDef.Id);

                } catch (er) {
                    LOG(`ERROR! Unable to assign 'Read' for Group ID: ${readerId}.`, er);
                }
            }

            await sleep(50);   // pause to let SPO do the work

        } catch (er) {
            LOG("ERROR! Unable to set list role assigment.", er);
        }
    }

    /** Resets and breaks item permissions. Sets Contribute role to users.
     * @param itemId List Item ID to reset and update unique permissions to.
     * @param userIds User IDs being granted Contributor permissions.
     */
    public async updateContentManagers(itemId: number, userIds?: number | number[]): Promise<void> {
        try {
            const item = this.api.items.getById(itemId),
                itemData = await item.select("Id", "ContentManagersId").get(),
                cmRole = await this.ensureRole("Content Manager");

            if (!userIds) userIds = itemData.ContentManagersId || [];
            if (typeof userIds === "number") userIds = [userIds];

            // reset inheritence - this should clear all existing users and resync with list permissions
            await item.resetRoleInheritance();
            await item.breakRoleInheritance(true);  // when breaking inheritance, copy list permissions

            // add users
            userIds.forEach(async id => {
                // add user permissions
                await item.roleAssignments.add(id, cmRole.Id);
            });

        } catch (er) {
            LOG(`ERROR! Unable to update Content Managers for item ID: ${ itemId}.`, er);
        }
    }
//#endregion

//#region USER
    public async getUserPermissions(): Promise<IUserPermissions> {
        const requests = [this.api.currentUserHasPermissions(PermissionKind.AddListItems),
            this.api.currentUserHasPermissions(PermissionKind.EditListItems),
            this.api.currentUserHasPermissions(PermissionKind.DeleteListItems)],
            [canAdd, canEdit, canDelete] = await Promise.all(requests);

        return {
            canAdd: canAdd,
            canEdit: canEdit,
            canDelete: canDelete
        };
    }

    public async getCurrentUser(): Promise<ISiteUserInfo> {
        return await sp.web.currentUser();
    }
//#endregion

//#region SEARCH
    /** Search this list for matching items.
     * @param queryText Text term to search for.
     * @param rowLimit Optional. Maximum number of results returned. Increasing this limit will increase the chances of 
     * returning items but decreases response times. Default: 100.
     * @returns Array of trimmed results. Trims results to list items.
     */
    public async search(queryText, rowLimit?: number): Promise<ISearchResult[]> {
        const path = location.protocol + "//" + location.host + this.properties.RootFolder.ServerRelativeUrl,
            results: ISearchResult[] = [],
            addedList = [],
            response = await sp.searchWithCaching(<ISearchQuery>{
                Querytext: `${escapeProperly(queryText)} path:${path}`,
                // Querytext: queryText,
                RowLimit: rowLimit || 100,
                EnablePhonetic: true,
                SourceId: this.properties.Id,
                SelectProperties: ["Title", "Path", "ListId"],
                EnableInterleaving: true
            });

        // trim results for this list only
        response.PrimarySearchResults.forEach(result => {
            const id = parseInt(result.Path.split("ID=").pop(), 10);
            if (!isNaN(id) && addedList.indexOf(id) === -1) {
                addedList.push(id);  // keep track of which items were added. avoid adding duplicate results
                results.push(result);
            }
        });

        LOG(`Search results for ${queryText}:`, response);

        return results;
    }
//#endregion

//#region UPLOAD
    public static async uploadFile(libraryRelativeUrl: string, file: File, folderName?: string): Promise<IItem> {
        let item: IItem = null;
        const folderRelativeUrl = [libraryRelativeUrl];


        try{
            // ensure folder exists
            if (folderName) {
                const folderExists = await DataFactory.ensureFolder(libraryRelativeUrl, folderName);
                if (folderExists) folderRelativeUrl.push(folderName);
            }

            // upload the file
            const fileAddResult = await sp.web.getFolderByServerRelativeUrl(folderRelativeUrl.join("/")).files.add(file.name, file, true);
            
            // get file/item details
            item = await fileAddResult.file.get();

        } catch (er) {
            LOG("ERROR! Unable to upload file:", file, er);
        }

        return item;
    }

    public static async ensureFolder(libraryRelativeUrl: string, folderName: string): Promise<boolean> {
        // get folder
        let folderInfo: IFolderInfo;
        
        try {
            folderInfo = await sp.web.getFolderByServerRelativeUrl([libraryRelativeUrl, folderName].join("/")).get();
        } catch (er) {/* silently */}

        if (!folderInfo || folderInfo.Exists === false) {
            try {
                // try to create the folder
                await sp.web.getFolderByServerRelativeUrl(libraryRelativeUrl).addSubFolderUsingPath(folderName);

            } catch (er) {
                LOG("ERROR! Unable to create subfolder:", libraryRelativeUrl, folderName, er);
                return false;
            }
        }

        return true;
    }
//#endregion

//#region HELPERS
    /** Creates the structure object from all list items.
     * @param webpartProperties Web Part Property object. Used to create the "Root" object
     * @returns Structure object
     */
    public async getDataStructure(homeTitle: string): Promise<Record<(number | string), IDataStructureItem>> {

        // get all items from the list
        const select = ["ID", "Title", "DisplayType", "Name", "OrgType", "OrderBy", "Parent/ID"], 
            expand = ["Parent"],
            orderBy: [string, boolean] = ["OrderBy", true],
            items = await this.getAllItems(select, expand, orderBy),
            orgTypeField = find(this.fields, ["InternalName", "OrgType"]),
            orgTypeOptions = (orgTypeField) ? orgTypeField.Choices : [],
            structure: Record<(number | string), IDataStructureItem> = {};

        // if no items returned (empty list?), return empty structure
        if (items.length === 0) return structure;

        // add root to structure
        structure["_root"] = {
            "ID": 0,
            "Title": homeTitle || "Home",
            "DisplayType": ["_root"],
            "OrderBy": -1000,
            "children": []
        };

        // add org types to structure
        orgTypeOptions.forEach((orgType, i) => {
            // create structure item
            structure[orgType] ={
                "ID": 0,
                "Title": orgType,
                "DisplayType": ["_orgType"],
                "OrderBy": -110 + i,
                "ParentID": "_root",
                "children": []
            };
        });

        // add all list items to stucture
        items.forEach(item => {
            structure[item.ID] = {
                "ID": item.ID,
                "Title": item.Title || "",
                "Name": item.Name || "",
                "DisplayType": item.DisplayType || [],
                "OrgType": item.OrgType,
                "OrderBy": item.OrderBy || 999999999,
                "ParentID": (item.Parent) ? item.Parent.ID : item.OrgType || null,
                "children": [],
                "data": null,
                "flags": {}
            };
        });

        // nest structure. add each item to it's parent.
        for (const key in structure) {
            if (Object.prototype.hasOwnProperty.call(structure, key)) {
                const item = structure[key];
                
                if (item.ParentID) {
                    const parent = structure[item.ParentID];
                    if (parent) parent.children.push(item);
                }
            }
        }

        // sort childen array
        // NOTE: do this after all children have been nested. avoid sorting multiple times
        const fnSort = (a: IDataStructureItem, b: IDataStructureItem) => { return a.OrderBy - b.OrderBy; };
        for (const key in structure) {
            if (Object.prototype.hasOwnProperty.call(structure, key)) {
                const item = structure[key];

                if (item.children.length > 1) item.children.sort(fnSort);
            }
        }

        return structure;
    }

    /** Get current web's site group information
     * @returns Array of site group inforamtion
     */
    public static async getSiteGroups(): Promise<ISiteGroupInfo[]> {
        try{

            const groups = await sp.web.siteGroups();
            return groups;

        } catch(er) {
            LOG("ERROR! Unable to get site groups:", er);
            return null;
        }
    }

    /** Get current web's site group information
     * @returns Array of site group inforamtion
     */
    public static async getSiteGroupMembers(groupID): Promise<ISiteUserInfo[]> {
        try{

            const users = await sp.web.siteGroups.getById(groupID).users();
            return users;

        } catch(er) {
            LOG("ERROR! Unable to get site group members:", er);
            return null;
        }
    }

    /**
     * Get list properties by list title.
     * @param title Title for list to get
     * @returns List properties. Returns null if not found.
     */
    private static async getListByTitle(title: string, ...select: string[]): Promise<IListInfo> {
        try {
            // will throw error if list doesn't exist
            const list = sp.web.lists.getByTitle(title);

            return await list
                .select.apply(list, select)
                .get({
                    "headers": { 
                        "Accept": "application/json;odata=nometadata",
                        "Cache-Control": "no-cache"
                    }
                });
        
        } catch(er) {
            LOG("ERROR! Unable to get all lists:", er);
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
        const list = await DataFactory.getListByTitle(title, "Title");
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
        ): Promise<Record<string, any>[]> {

        const apiLists = sp.web.lists;

        const lists = await apiLists
            .select.apply(apiLists, select)
            .expand.apply(apiLists, expand)
            .filter(filter)
            .get({
                "headers": {"Accept": "application/json;odata=nometadata"}
            });

        return lists;
    }

    /**
     * REST request to retrieve all SP lists from site.
     * @param absoluteWebUrl Full URL to web
     * @returns Array of list objects.
     */
     public static async getAllLibraries(absoluteWebUrl: string): Promise<IDocumentLibraryInformation[]> {

        return await sp.site.getDocumentLibraries(absoluteWebUrl);
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
            return { "valid": false, "message": "List name cannot contain special characters.\nUnderscores, parathisis' \
                and dashes are allowed." };
        }
        
        // if list exists = not valid;
        if (existingListNames.indexOf(name) > -1) {
            return { "valid": false, "message": "List already exists!\nEnter a new list name\nor select from the list dropdown." };
        }

        return { "valid": true, "message": "" };
    }
//#endregion
}