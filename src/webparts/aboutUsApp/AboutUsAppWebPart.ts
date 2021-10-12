import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConditionalGroup,
    IPropertyPaneConfiguration,
    IPropertyPaneDropdownOption,
    IPropertyPaneField,
    IPropertyPaneGroup,
    IPropertyPanePage,
    PropertyPaneButton,
    PropertyPaneButtonType,
    PropertyPaneCheckbox,
    PropertyPaneDropdown,
    PropertyPaneDropdownOptionType,
    PropertyPaneHorizontalRule,
    PropertyPaneLabel,
    PropertyPaneLink,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AboutUsAppWebPartStrings';
import AboutUsApp, { IAboutUsAppProps } from './components/AboutUsApp';

import { SPComponentLoader } from "@microsoft/sp-loader";
import DataFactory, { IListValidationResults, TAboutUsRoleDef } from './components/DataFactory';
import CustomDialog from './components/CustomDialog';
import { trim, escape, find } from 'lodash';
import { PropertyPaneDescription } from 'AboutUsAppWebPartStrings';
import { GroupShowAll, textAreaProperties } from 'office-ui-fabric-react';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';

// declare global {
//     interface Window {
//         formValue?: any;
//     }
// }
export interface IAboutUsAppFieldOption {
    required: boolean;
    controlled: boolean;
}
export interface IAboutUsAppWebPartProps {
    description: string;
    displayType: string;
    displayTypeOptions: Partial<IPropertyPaneDropdownOption[]>;
    listName: string;
    ppListName_dropdown: string | number;
    orgchart_key: { [color: string]: string };
    fields: { [fieldName: string]: IAboutUsAppFieldOption };
    ownerGroup: number;
    managerGroup: number;
    readerGroup: number;
}


export default class AboutUsAppWebPart extends BaseClientSideWebPart<IAboutUsAppWebPartProps> {
//#region PROPERTIES
    private list_: DataFactory;     // Content List object with methods and helpers

    private propertyPane_ = {
        "defaultName": "AboutUs_Content",           // suggested list name. may not be the list this app is using.
        "isReady": false,   // has data and propertypane is ready
        "showLoading": true,    // show loading icon
        "showPickListMenu": true,
        "propertyChanged": false,   // flag when any property pane field changes. onClose, updates render if anything changed
        "data": {
            "optListNames": [],
            "optSiteGroups": [],
            "optOrgChartColors": []
        }
    };

    // Processing... modal message properties.
    private modalMsg_processing = [
        '**WARNING** DO NOT CLOSE THIS TAB!\n\nThis message will disappear when the process has been completed.',
        'Processing...'
    ];
//#endregion

//#region PRE-RENDER
    public constructor() {
        super();
    }

    protected async onInit(): Promise<void> {
        return await super.onInit().then(async () => {
            // web part context is only set for the base class. 
            // send the context to the other classes that need them.
            AboutUsApp.ctx = this.context;

            // create the 'About-Us' list (datafactory), SP-PnP requires the web part context.
            this.list_ = new DataFactory(this.context);

            // check if list exists; may have gotten deleted, renamed, or never initialized
            // in the network console, this request may return a 404 status if the list doesn't exist.
            if (typeof this.properties.listName === "string" && this.properties.listName.length > 0) {
                // list name was previously populated, ensure list exists
                try {
                    await this.list_.ensureList(this.properties.listName);
                    this.propertyPane_.showPickListMenu = !this.list_.exists;

                } catch (er) {
                    AboutUsAppWebPart.DEBUG(`ERROR! Could not ensure '${ this.properties.listName }' list properties or fields. ` +
                        `Check to see if the list exists and you have the proper permissions.`, er);
                }

            } else {
                // list doesn't exist.
                await this.updateRenderProperty("listName", "");
                
            }

        });
    }

//#endregion

//#region RENDER
    public async render(): Promise<void> {
        const ELEMENT: React.ReactElement<IAboutUsAppProps> = React.createElement(
            AboutUsApp,
            {
                displayType: this.properties.displayType,
                webpart: this.properties,
                list: this.list_
            }
        );

        ReactDom.render(ELEMENT, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    // @ts-ignore
    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    /**
     * Update properties that affect the display.
     * @param properties Key/value pair of property(s) that changed. Properties that affects the display.
     */
    private async updateRenderProperty(propertyName?: "listName" | "displayType", value?: any, renderNow: boolean = true): Promise<void> {
        
        switch (propertyName) {
        
            case "listName":
                this.properties.listName = <string>value;
                break;

            case "displayType":
                this.properties.displayType = <string>value;
                break;

        }
        
        // re-render
        if (this.renderedOnce && renderNow) await this.render();
    }
//#endregion

//#region PROPERTY PANE
    /** Property Pane Configuration
     * @returns SPFx Property pane configuration
     */
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        const _this = this,
            pages : IPropertyPanePage[] = [];

        // is property pane ready - has data
        if (this.propertyPane_.isReady === false) {
            this.propertyPane_.showLoading = true;
            
            // get data
            Promise.all([this.populateListNameOptions(), this.populateSiteGroupOptions()]).then(responses => {
                _this.propertyPane_.showLoading = false;
                _this.propertyPane_.isReady = true;
                _this.context.propertyPane.refresh();
            });

        } else {    // got data, propertypane is ready

            // General page
            pages.push(this.propertyPanePage_General());

            // other pages will only display after a list has been selected
            if (this.list_.exists) {
                // other pages are based on which display type is selected
                switch (this.properties.displayType) {
                    case "page":
                        pages.push(this.propertyPanePage_Permissions());
                        pages.push(this.propertyPanePage_Fields());
                        break;
                
                    case "orgchart":
                        pages.push(this.propertyPanePage_OrgChart()); 
                        break;
                
                    case "accordian":
                
                        break;
                
                    case "phone":
            
                        break;
                
                    case "datatable":
        
                        break;
                                
                    case "leadershipbroadcast":
                        
                        break;
                
                    default:
                        break;
                }
            }

        }

        return {
            showLoadingIndicator: this.propertyPane_.showLoading,
            pages: pages
        };
    }

    /** PropertyPane Page
     * @returns PropertyPane page and fields.
     */
    private propertyPanePage_General(): IPropertyPanePage {
        const group_General = {
                groupName: "General",
                groupFields: []
            },
            group_List = {
                groupName: "Content list",
                groupFields: []
            },
            group_Back = {
                groupName: "",
                groupFields: []
            };

        // display type
        group_General.groupFields.push(
            PropertyPaneDropdown("displayType", {
                label: "Display Type",
                options: this.properties.displayTypeOptions
            })
        );

        // list menu
        if (this.list_.exists && this.propertyPane_.showPickListMenu === false) {
            // show update list menu

            // link: current list
            group_List.groupFields.push(
                PropertyPaneLink("lnkListName", {
                    text: "Current list: " + this.properties.listName,
                    href: this.list_.properties.RootFolder.ServerRelativeUrl,
                    target: "_blank"
                })
            );

            // label: Update list
            group_List.groupFields.push(
                PropertyPaneLabel("lblUpdateList", {
                    text: "Update list properties & fields:"
                })
            );

            // button: Update list
            group_List.groupFields.push(
                PropertyPaneButton("btnUpdateList", {
                    text: "Update now!",
                    onClick: async () => {
                        const modalMsg = CustomDialog.modalMsg.apply(null, this.modalMsg_processing);
                        try {
                            await this.list_.ensureList(this.list_.title, true, true, true);
                            await this.resetListPermissions();
                            modalMsg.close();

                        } catch (er) {
                            modalMsg.close();
                            AboutUsAppWebPart.DEBUG(`ERROR! Unable to update '${ this.properties.listName }' list with About-Us properties & fields.`, er);
                            await CustomDialog.alert("Something went wrong! See the console for details.", "ERROR");

                        }

                        this.propertyPane_.showPickListMenu = false;
                        this.context.propertyPane.refresh();
                    }
                })
            );

            // label: Select a list
            group_List.groupFields.push(
                PropertyPaneLabel("lblUpdateDescription", {
                    text: "Updating the list will try to update the current list with properties from the About-Us list template."
                })
            );

            // button: Select a list
            group_Back.groupName = "Select or create a new list";
            group_Back.groupFields.push(
                PropertyPaneButton("btnBackToSelectNewList", {
                    buttonType: PropertyPaneButtonType.Icon,
                    icon: "Back",
                    text: "Back",
                    onClick: () => {
                        this.propertyPane_.showPickListMenu = true;
                        this.context.propertyPane.refresh();
                    }
                })
            );

        } else {    // show pick-a-list menu

            // button: Create List
            group_List.groupFields.push(
                PropertyPaneButton("btnCreateNewList", {
                    buttonType: PropertyPaneButtonType.Normal,
                    text: "Create New List",
                    onClick: this.createList_click.bind(this)
                })
            );

            //label: or
            group_List.groupFields.push(
                PropertyPaneLabel("lblOr", {
                    text: " or"
                })
            );

            // dropdown: Select from existing list
            group_List.groupFields.push(
                PropertyPaneDropdown("ppListName_dropdown", {
                    label: 'Select from an existing list:',
                    options: this.propertyPane_.data.optListNames
                })
            );

            if (this.list_.exists) {
                group_Back.groupName = "Back to Update List";
                group_Back.groupFields.push(
                    PropertyPaneButton("btnBackToSelectNewList", {
                        buttonType: PropertyPaneButtonType.Icon,
                        icon: "Back",
                        text: "Back",
                        onClick: () => {
                            this.propertyPane_.showPickListMenu = false;
                            this.context.propertyPane.refresh();
                        }
                    })
                );
            }
        }

        return {
            header: {
                description: this.properties.description
            },
            groups: [group_General, group_List, group_Back]
        };
    }

    /** PropertyPane Page
     * @returns PropertyPane page and fields.
     */
    private propertyPanePage_Permissions(): IPropertyPanePage {
        const group_Permissions = {
                groupName: "Edit Permissions",
                groupFields: []
            },
            group_PermissionLevelLink = {
                groupName: "Manage Site Permission Levels",
                groupFields: []
            };

        // try to prefill empty (new) roles
        this.setExistingRoles.call(this);

        
        // label: Permission level label
        group_PermissionLevelLink.groupFields.push(
            PropertyPaneLabel("lblPermissionLevel", {
                text: "The About-Us list does not inherit permissions. This allows the " +
                    "About-Us organizational structure (hierarchy) to be managed by a different " +
                    "group of people, like the organization's manpower group."
            })
        );
        // link: Site Permission Level
        group_PermissionLevelLink.groupFields.push(
            PropertyPaneLink("lnkPermissionLevel", {
                text: "Open permission level settings",
                href: this.context.pageContext.web.serverRelativeUrl + "/_layouts/15/user.aspx",
                target: "_blank"
            })
        );


        // Owner Permission Group
        group_Permissions.groupFields.push(
            PropertyPaneDropdown("ownerGroup", {
                label: "Owners Group",
                options: this.propertyPane_.data.optSiteGroups
            })
        );
        // owner group description
        group_Permissions.groupFields.push(
            PropertyPaneLabel("lblOwnerGroup", {
                text: "Owners have full control (ownership) of the About-Us information. " +
                    "This is normally the site's Owners group."
            })
        );

        // Manager Permission Group
        group_Permissions.groupFields.push(
            PropertyPaneDropdown("managerGroup", {
                label: "About-Us Manager Group",
                options: this.propertyPane_.data.optSiteGroups
            })
        );
        // manager group description
        group_Permissions.groupFields.push(
            PropertyPaneLabel("lblManagerGroup", {
                text: "About-Us Managers can add, edit all fields including controlled fields, " +
                    "modify the organizational structure, & delete About-Us entries. " +
                    "This group should have at least Read rights to the site and will be " +
                    "granted Contributor rights to the About-Us list."
            })
        );

        // Reader Permission Group
        group_Permissions.groupFields.push(
            PropertyPaneDropdown("readerGroup", {
                label: "Readers Group (Visitors)",
                options: this.propertyPane_.data.optSiteGroups
            })
        );
        // reader group description
        group_Permissions.groupFields.push(
            PropertyPaneLabel("lblReaderGroup", {
                text: "Reader group members can view About-Us information. " +
                    "Normally the site's Visitors group."
            })
        );


        return {
            header: {
                description: "List permissions settings. Modifying permissions resets the " +
                    "list permission settings and updates the permissions for every item in the list."
            },
            groups: [group_PermissionLevelLink, group_Permissions]
        };
    }

    /** PropertyPane Page
     * @returns PropertyPane page and fields.
     */
    private propertyPanePage_Fields(): IPropertyPanePage {
        const groups = [];

        groups.push({
            groupName: "About-Us Fields",
            groupFields: [
                PropertyPaneLabel("lblRequiredDesc", {
                    text: "Required: These fields are required to be filled out. If disabled (grayed-out), this field is required by the app and cannot be changed."
                }),
                PropertyPaneLabel("lblControlledFieldDesc", {
                    text: "Controlled Field: These fields can only be updated by members of the About-Us Owners or Managers group."
                })
            ]
        });

        this.list_.fields.forEach(field => {
            if (!(field.InternalName in this.properties.fields)) this.properties.fields[field.InternalName] = {
                required: field.Required || false,
                controlled: false
            };

            groups.push({
                groupName: `${field.Title}:`,
                groupFields: [
                    PropertyPaneCheckbox(`fields.${field.InternalName}.required`, {
                        text: "Required",
                        disabled: field.Required || false,
                        checked: this.properties.fields[field.InternalName].required
                    }),
                    PropertyPaneCheckbox(`fields.${field.InternalName}.controlled`, {
                        text: "Controlled Field",
                        checked: this.properties.fields[field.InternalName].controlled
                    }),
                    PropertyPaneHorizontalRule()
                ]
            });
        });

        return {
            header: {
                description: "Select which fields should be mandatory and which fields are only editable by the About-Us Managers."
            },
            groups: groups
        };
    }

    /** PropertyPane Page
     * @returns PropertyPane page and fields.
     */
    private propertyPanePage_OrgChart(): IPropertyPanePage {
        const group_Colors = {
            groupName: "Org-Chart Color Meanings:",
            groupFields: []
        },
            fieldColors = this.populateOrgChartColorOptions();

        // Org Chart: show if display type is OrgChart & content list was ensured
        if (fieldColors.length > 0) {
            // label: Key
            group_Colors.groupFields.push(
                PropertyPaneLabel("lblOrgChartColor", {
                    text: "Legend:"
                })
            );

            // create each color label & field
            fieldColors.forEach( color => {

                // ensure the key exists in web part properties
                if (!(color.key in this.properties.orgchart_key)) {
                    this.properties.orgchart_key[color.key] = "";
                }
                // textbox: [COLOR-KEY]
                group_Colors.groupFields.push(
                    PropertyPaneTextField(`orgchart_key.${ color.key }`, {
                        label: `${ color.text }:`
                    })
                );
            });

        }

        return {
            header: {
                description: "Org-Chart settings."
            },
            groups: [group_Colors]
        };
    }

    /** SPFx PropertyPane onChange event handler. Triggers anytime a PropertyPane form element is changed.
     * @param propertyPath PropertyPane's target name. Use this property to determine which element triggered the onChange.
     * @param oldValue Previous value
     * @param newValue New value
     */
    public async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
        this.propertyPane_.propertyChanged = true;

        // do something based on which property (propertyPath) changed
        switch (propertyPath) {
            case "displayType": // display type dropdown changed
                this.updateRenderProperty("displayType", newValue);
                this.context.propertyPane.refresh();
                break;

            case "ppListName_dropdown":  // select a list dropdown changed
                
                if (newValue === "") return;

                // this.propertyPane_.showLoading = true;
                // this.context.propertyPane.refresh();

                // show warning. we need to ensure the 'About-Us' list properties and fields are set/available
                CustomDialog.confirm(
                    `Do you want to modify "${newValue}" list with About-Us fields and properties? Click "Continue" to proceed.`,
                    "IMPORTANT! Changes may be irreversible.",
                    {
                        "yes": "Continue",
                        "no": "Cancel"
                    }).then( response => {

                        // 'continue'?
                        if (response === true) {
                            const modalMsg = CustomDialog.modalMsg.apply(null, this.modalMsg_processing);
                            this.list_.ensureList(newValue).then( () => {
                                this.updateRenderProperty("listName", newValue);
                                this.properties.ppListName_dropdown = newValue;
                                this.propertyPane_.showPickListMenu = false;
                                this.context.propertyPane.refresh();
                                modalMsg.close();

                            }).catch( er => {
                                modalMsg.close();
                                AboutUsAppWebPart.DEBUG(`ERROR! Unable to check or update '${ newValue }' list with About-Us properties & fields.`, er);
                                CustomDialog.alert("Something went wrong! See the console for details.", "ERROR").then( () => {
                                    this.properties.ppListName_dropdown = "";
                                    this.propertyPane_.showPickListMenu = true;
                                    this.context.propertyPane.refresh();
                                });

                            });

                        } else {
                            // cancelled
                            this.properties.ppListName_dropdown = "";
                            this.propertyPane_.showPickListMenu = true;
                            this.context.propertyPane.refresh();

                        }
                    });

                break;
        
            case "ownerGroup": // list permissions updated
                this.properties.ownerGroup = newValue;
                this.resetListPermissions();
                break;

            case "managerGroup": // list permissions updated
                this.properties.managerGroup = newValue;
                this.resetListPermissions();
                break;


            case "readerGroup": // list permissions updated
                this.properties.readerGroup = newValue;
                this.resetListPermissions();
                break;

            default:
                AboutUsAppWebPart.DEBUG("Uncaught PropertyPane change handler:", propertyPath, oldValue, newValue);
        }
    }

    /** SPFx onClose event handler. Refresh render if any property pane field changed. */
    protected async onPropertyPaneConfigurationComplete(): Promise<void> {
        if (this.propertyPane_.propertyChanged) {
            this.propertyPane_.propertyChanged = false;
            await this.updateRenderProperty();
        }
    }

    /** 'Create New List' click event handler. Prompts user for a list name, then creates it.
     * @param evt Click event object
     * @returns Promise. Resolved when the dialog closes and the list is created
     */
     private async createList_click(evt: Event): Promise<void> {
        let newListName: string = this.properties.listName || this.propertyPane_.defaultName;
        let validation: IListValidationResults = { "valid": false, "message": "" };
        const existingListNames = this.propertyPane_.data.optListNames.map(opt => opt.text);

        // update property pane view
        this.propertyPane_.showLoading = true;
        this.context.propertyPane.refresh();
        
        // show form that allows the user to enter a list name.
        // keep showing unless user exits or enters a valid list name.
        do {
            newListName = await CustomDialog.prompt(
                "Enter new list name:",
                "Create New List",
                {
                    "value": newListName,
                    "description": "Must be unique and no special characters.",
                    "error": validation.message
                }
            );

            // exit if user clicked "Cancel" or closed the dialog
            if (newListName === null) {
                this.propertyPane_.showPickListMenu = true;
                this.propertyPane_.showLoading = false;
                this.context.propertyPane.refresh();
                return;
            }

            // clean up (trim) use input
            newListName = trim(newListName);
        } while ((validation = DataFactory.validateListName(newListName, existingListNames)).valid === false && newListName !== null);

        // try to create the list and fields
        const modalMsg = CustomDialog.modalMsg.apply(null, this.modalMsg_processing);
        try {
            await this.list_.ensureList(newListName, true, true, true, true);
            this.setExistingRoles();
            await this.updateRenderProperty("listName", newListName);
            this.propertyPane_.showPickListMenu = false;    // show update list menu

        } catch (er) {
            AboutUsAppWebPart.DEBUG(`ERROR! Unable to create '${ newListName }' list.`, er);
            await CustomDialog.alert("Something went wrong! See the console for details.", "ERROR");
            this.propertyPane_.showPickListMenu = true;
            
        }
        modalMsg.close();

        // refresh property pane
        this.propertyPane_.showLoading = false;
        this.context.propertyPane.refresh();

    }

    /** Tries to set the RoleDefs (Owners, Managers, Visitors) property settings. */
    private setExistingRoles() {
        const groups: {name: string, role: TAboutUsRoleDef}[] = [
                { name: "ownerGroup", role: "Full Control" },
                { name: "managerGroup", role: "Contribute" },
                { name: "readerGroup", role: "Read" }
            ],
            _getRoleAssignmentsFor = (role: TAboutUsRoleDef): number => {
                if ("RoleAssignments" in this.list_.properties && "results" in this.list_.properties.RoleAssignments) {
                    const roles = <any[]>this.list_.properties.RoleAssignments.results;

                    // loop through each role, look for matching role. return first match
                    for (let i = 0; i < roles.length; i++) {
                        const roleDef = roles[i],
                            bindings = <any[]>roleDef.RoleDefinitionBindings.results,
                            memberId = <number>roleDef.Member.Id;

                        // loop through each of the bindings, look for a matching role
                        for (let ii = 0; ii < bindings.length; ii++) {
                            const binding = bindings[ii];
                            
                            if (binding.Name === role) return memberId;
                        }
                    }
                }

                return null;
            };

        groups.forEach( function (group) {
            // try to update the group if not already set
            if (this.properties[group.name] === null) {
                // get the owners (Full Control) from the list's role assignments
                const roleAssignment = _getRoleAssignmentsFor.call(this, group.role);
                if (roleAssignment) this.properties[group.name] = roleAssignment;
                
            }
        }.bind(this));
    }

    /** Resets and updated list's unique permissions, then resyncs permissions for all items. */
    private async resetListPermissions(): Promise<void> {
        // show modal message
        const modalMsg = CustomDialog.modalMsg.apply(null, this.modalMsg_processing);

        try {
            // reset list permissions
            await this.list_.resetListRoleAssignments(this.properties.ownerGroup, this.properties.managerGroup, this.properties.readerGroup);
            
            // then we need to update content managers for every item
            await this.updateContentManagersForAllItems();

        } catch (er) {
            AboutUsAppWebPart.DEBUG("ERROR! Unable to reset list permissions.", er);
            CustomDialog.alert("Something went wrong. Open the console to see error details.");
        }

        modalMsg.close();
    }

    /** Updates Content Managers for all list items */
    private async updateContentManagersForAllItems(): Promise<void> {
        try {
            const items = await this.list_.api.items.select("ID").getAll();

            items.forEach(async item => {
                await this.list_.updateContentManagers(item.Id);
            });
        } catch (er) {
            AboutUsAppWebPart.DEBUG("ERROR! Unable to update content managers for all items.", er);
        }
    }
//#endregion

//#region HELPERS
    /** Gets and populates list name options 
     * @returns Array of existing list names. Formatted: [{key: string, text: string}, ...];
     */
    private async populateListNameOptions(): Promise<IPropertyPaneDropdownOption[]> {
        try {
            if (this.propertyPane_.data.optListNames.length === 0) {
                const DATA = await DataFactory.getAllLists(),
                    listNames = DATA.map(i => i.Title);

                listNames.sort();
                this.propertyPane_.data.optListNames = listNames.map(name => { return {key: name, text: name}; });

                // add a blank option
                this.propertyPane_.data.optListNames.unshift({key: "", text: "-"});
                
            }
        } catch (er) {
            AboutUsAppWebPart.DEBUG("ERROR! Unable to get list names.", er);
        }

        return this.propertyPane_.data.optListNames;
    }

    /** Gets and populates site group options 
     * @returns Array of existing site groups. Formatted: [{key: string, text: string}, ...];
     */
     private async populateSiteGroupOptions(): Promise<IPropertyPaneDropdownOption[]> {
        try {
            if (this.propertyPane_.data.optSiteGroups.length === 0) {
                const groups = await DataFactory.getSiteGroups();

                this.propertyPane_.data.optSiteGroups = [];
                groups.forEach(group => {
                    if (group.OwnerTitle !== "System Account") {
                        this.propertyPane_.data.optSiteGroups.push({ key: group.Id, text: group.Title });
                    }
                });
            }

        } catch (er) {
            AboutUsAppWebPart.DEBUG("ERROR! Unable to get site groups.", er);
        }

        return this.propertyPane_.data.optSiteGroups;
    }

    /** Gets and populates org chart color options 
     * @returns Array of existing org chart colors. Formatted: [{key: string, text: string}, ...];
     */
     private populateOrgChartColorOptions(): IPropertyPaneDropdownOption[] {
        try {
            if (this.propertyPane_.data.optOrgChartColors.length === 0 && this.list_.exists) {
                const fieldValue = find(this.list_.fields, ["InternalName", "OrgChartColor"]);
                
                this.propertyPane_.data.optOrgChartColors = fieldValue.Choices.map(val => { 
                    const color = val.split(/ ?\(/)[0],
                        key = trim(color).replace(/\s/g, "").toLowerCase();

                    return {key: color, text: color};
                });

            }
        } catch (er) {
            AboutUsAppWebPart.DEBUG("ERROR! Unable to get Org Chart Colors.", er);
        }

        return this.propertyPane_.data.optOrgChartColors;
    }

    /** Prints our debug messages. Decorated console.info() or console.error() method.
     * @param args Message or object to view in the console. If message starts with "ERROR", DEBUG will use console.error().
     */
    private static DEBUG(...args: any[]) {
        // is an error message, if first argument is a string and contains "error" string.
        const isError = (args.length > 0 && (typeof args[0] === "string")) ? args[0].toLowerCase().indexOf("error") > -1 : false;
        args = ["(About-Us)"].concat(args);

        if (window && window.console) {
            if (isError && console.error) {
                console.error.apply(null, args);

            } else if (console.info) {
                console.info.apply(null, args);

            }
        }
    }
//#endregion
}
