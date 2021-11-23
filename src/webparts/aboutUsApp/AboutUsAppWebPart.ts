import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConditionalGroup,
    IPropertyPaneConfiguration,
    IPropertyPaneDropdownOption,
    IPropertyPaneGroup,
    IPropertyPanePage,
    PropertyPaneButton,
    PropertyPaneButtonType,
    PropertyPaneCheckbox,
    PropertyPaneChoiceGroup,
    PropertyPaneDropdown,
    PropertyPaneDropdownOptionType,
    PropertyPaneHorizontalRule,
    PropertyPaneLabel,
    PropertyPaneLink,
    PropertyPaneTextField,
    PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AboutUsAppWebPartStrings';
import AboutUsApp, { IAboutUsAppProps } from './components/AboutUsApp';

import { SPComponentLoader } from "@microsoft/sp-loader";
import DataFactory, { IListValidationResults, TAboutUsRoleDef } from './components/DataFactory';
import CustomDialog from './components/CustomDialog';
import { trim, escape, find, trimEnd } from 'lodash';

//#region INTERFACES, TYPES & ENUMS
export interface IAboutUsAppFieldOption {
    required: boolean;
    controlled: boolean;
}
export interface IAboutUsAppWebPartProps {
    urlParam: string;
    description: string;
    displayType: string;
    displayTypeOptions: Partial<IPropertyPaneDropdownOption[]>;
    listName: string;
    ppListName_dropdown: string | number;
    homeTitle: string;
    logo: IFilePickerResult;
    startingID: number;
    showTaskAuth: boolean;
    validateEvery: number;
    externalRepo: string;
    appMessage: string;
    appMessageIsAlert: boolean;
    pageTemplate: string;
    orgchart_key: Record<string, string>;
    fields: Record<string, IAboutUsAppFieldOption>;
    ownerGroup: number;
    managerGroup: number;
    readerGroup: number;
    broadcastDays: number;
}
//#endregion


export default class AboutUsAppWebPart extends BaseClientSideWebPart<IAboutUsAppWebPartProps> {
//#region PROPERTIES
    private list_: DataFactory;     // Content List object with methods and helpers

    private propertyPane_ = {
        "defaultName": "AboutUs_Content",           // suggested list name. may not be the list this app is using.
        "isReady": false,   // has data and propertypane is ready
        "showLoading": true,    // show loading icon
        "showPickListMenu": true,
        "propertyChanged": false,   // flag when any property pane field changes. onClose, updates render
        "data": {
            "optListNames": [],
            "optLibraryNames": [],
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

//#region RENDER
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
                    LOG(`ERROR! Could not ensure '${ this.properties.listName }' list properties or fields. ` +
                        `Check to see if the list exists and you have the proper permissions.`, er);
                }

            } else {
                // list doesn't exist.
                await this.updateRenderProperty("listName", "");
                
            }

        });
    }

    public async render(): Promise<void> {
        const element: React.ReactElement<IAboutUsAppProps> = React.createElement(AboutUsApp, {
                displayType: this.properties.displayType,
                properties: this.properties,
                list: this.list_
            }
        );

        ReactDom.render(element, this.domElement);
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
            
            // get all data
            Promise.all([
                this.populateListNameOptions(),
                this.populateLibraryNameOptions(),
                this.populateSiteGroupOptions()
            ]).then(responses => {

                // refresh property pane
                _this.propertyPane_.showLoading = false;
                _this.propertyPane_.isReady = true;
                _this.context.propertyPane.refresh();
            });

        } else {    // got data, propertypane is ready


            // other pages will only display after a list has been selected
            if (this.list_.exists) {
                // General page. settings that affect multiple views
                pages.push(this.propertyPanePage_General());

                // other pages are based on which display type is selected
                switch (this.properties.displayType) {
                    case "page":
                        pages[0].groups = pages[0].groups.concat(this.propertyPaneGroups_Page());
                        pages.push(this.propertyPanePage_Permissions());
                        pages.push(this.propertyPanePage_Fields());
                        break;
                
                    case "orgchart":
                        pages[0].groups = pages[0].groups.concat(this.propertyPaneGroups_OrgChart()); 
                        break;
                
                    case "accordian":
                
                        break;
                
                    case "phone":
            
                        break;
                
                    case "datatable":
        
                        break;
                                
                    case "broadcast":
                        pages[0].groups = pages[0].groups.concat(this.propertyPaneGroups_Broadcast());
                        break;
                
                    default:
                        break;
                }
            }

            pages.push(this.propertyPanePage_List());
        }

        return {
            showLoadingIndicator: this.propertyPane_.showLoading,
            pages: pages
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

                // show warning. we need to ensure the 'About-Us' list properties and fields are set/available
                CustomDialog.confirm(
                    'Check and update the list to ensure it has the required list fields and properties.\
                    \nClick "Continue" to proceed?',
                    `"${newValue}" list: Update List Properties`,
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
                                LOG(`ERROR! Unable to check or update '${ newValue }' list with About-Us properties & fields.`, er);
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
                // LOG("Uncaught PropertyPane change handler:", propertyPath, oldValue, newValue);
        }
    }

    /** SPFx onClose event handler. Refresh render if any property pane field changed. */
    protected async onPropertyPaneConfigurationComplete(): Promise<void> {
        if (this.propertyPane_.propertyChanged) {
            this.propertyPane_.propertyChanged = false;
            await this.updateRenderProperty();
        }
    }

    // PAGES & GROUPS
    /** PropertyPane General - Settings that affect multiple views
     * @returns PropertyPane page and fields.
     */
    private propertyPanePage_General(): IPropertyPanePage {
        const group = {
                groupName: "",
                groupFields: []
            },
            group_General = {
                groupName: "General",
                groupFields: []
            };

        // display type
        group.groupFields.push(
            PropertyPaneDropdown("displayType", {
                label: "Display web part as",
                options: this.properties.displayTypeOptions
            })
        );

        // Home name
        group_General.groupFields.push(
            PropertyPaneTextField("homeTitle", {
                label: "Edit root breadcrumb text: (Default: Home)"
            })
        );

        // Starting ID
        group_General.groupFields.push(
            PropertyPaneLabel("lblStartingID", {
                text: "On first render, this About-Us item will be displayed by default. \
                    If ID is 0 (zero) or the item ID does not exist, \
                    the default starting item will be the item with the lowest ID (created first/earliest)."
            })
        );
        group_General.groupFields.push(
            PropertyFieldNumber("startingID", {
                key: "startingID",
                label: "ID for the default starting item",
                description: "Numbers only. If '0' (zero), the web part will display the first availble item by default.",
                value: this.properties.startingID
            })
        );

        return {
            header: {
                description: this.properties.description
            },
            groups: [group, group_General]
        };
    }

    /** PropertyPane Groups
     * @returns PropertyPane groups array.
     */
     private propertyPaneGroups_Page(): (IPropertyPaneGroup | IPropertyPaneConditionalGroup)[] {
        const group_logo = {
                groupName: "Default Logo",
                groupFields: []
            },
            group = {
                groupName: "About-Us Page Settings",
                groupFields: []
            },
            group_Message = {
                groupName: "Page Banner Settings",
                groupFields: []
            };

        // Default logo
        group_logo.groupFields.push(
            PropertyFieldFilePicker("logo", {
                properties: this.properties,
                key: "logo",
                context: this.context,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                onSave: (e: IFilePickerResult) => { this.properties.logo = e; },
                label: "Default logo: Max resolution = 100px x 100px. Allowed: JPG, JPEG, or PNG",
                buttonLabel: "Select an image",
                filePickerResult: this.properties.logo,
                accepts: ["jpg", "jpeg", "png", "svg", "ico"],
                hideWebSearchTab: true,
                hideStockImages: true,
                hideOrganisationalAssetTab: true,
                hideOneDriveTab: true,
                hideLocalUploadTab: true
            })
        );

        // Display Tasking Authority
        group.groupFields.push(
            PropertyPaneToggle("showTaskAuth", {
                label: "Display Tasking Authority"
            })
        );

        // Validate Information Every...
        group.groupFields.push(
            PropertyPaneDropdown("validateEvery", {
                label: "About-Us information is valid for",
                options: [
                    {key: -1, text: "Never expires"},
                    {key: 30, text: "30 days"},
                    {key: 90, text: "90 days"},
                    {key: 180, text: "180 days"},
                    {key: 365, text: "1 year"},
                    {key: 730, text: "2 years"}
                ]
            })
        );

        // bio/images/logo repo location
        group.groupFields.push(
            PropertyPaneLabel("lblExternalRepo", {
                text: "An external repository (Document Library) is useful for content managers to upload bios and images \
                    for each of the different offices (branches, sections, divisions, directorates, orgs...). \
                    The minimum permissions should allow all visitors to view the files and \
                    allow content managers to upload/add, edit, & delete files as needed."
            })
        );
        group.groupFields.push(
            PropertyPaneDropdown("externalRepo", {
                label: 'Select external repository (Document Library w/ proper permission) for bios, bio images and logos:',
                options: this.propertyPane_.data.optLibraryNames
            })
        );

        // App message
        group_Message.groupFields.push(
            PropertyPaneLabel("lblAppMessage", {
                text: "Display important notifications at the top of this web part. \
                    Highlighting the notification will decorate the message as important."
            })
        );
        group_Message.groupFields.push(
            PropertyPaneTextField("appMessage", {
                label: "Enter notification: If blank, the notification section doesn't display."
            })
        );
        group_Message.groupFields.push(
            PropertyPaneToggle("appMessageIsAlert", {
                label: "Highlight notification"
            })
        );

        return [group_logo, group, group_Message];
    }
    /** PropertyPane Page
     * @returns PropertyPane page and fields.
     */
     private propertyPanePage_List(): IPropertyPanePage {
        const group = {
                groupName: "",
                groupFields: []
            },
            group_Back = {
                groupName: "",
                groupFields: []
            };

        // list menu
        if (this.list_.exists && this.propertyPane_.showPickListMenu === false) {
            // show update list menu
            group.groupName = "Update List";

            // link: current list
            group.groupFields.push(
                PropertyPaneLink("lnkListName", {
                    text: "Current list: " + this.properties.listName,
                    href: this.list_.properties.RootFolder.ServerRelativeUrl,
                    target: "_blank"
                })
            );

            // label: Update list
            group.groupFields.push(
                PropertyPaneLabel("lblUpdateList", {
                    text: "Update list properties & fields:"
                })
            );

            // button: Update list
            group.groupFields.push(
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
                            LOG(`ERROR! Unable to update '${ this.properties.listName }' list with About-Us properties & fields.`, er);
                            await CustomDialog.alert("Something went wrong! See the console for details.", "ERROR");

                        }

                        this.propertyPane_.showPickListMenu = false;
                        this.context.propertyPane.refresh();
                    }
                })
            );

            // label: Select a list
            group.groupFields.push(
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
            group.groupName = "Initialize or Select List";

            // button: Create List
            group.groupFields.push(
                PropertyPaneButton("btnCreateNewList", {
                    buttonType: PropertyPaneButtonType.Normal,
                    text: "Create New List",
                    onClick: this.createList_click.bind(this)
                })
            );

            //label: or
            group.groupFields.push(
                PropertyPaneLabel("lblOr", {
                    text: " or"
                })
            );

            // dropdown: Select from existing list
            group.groupFields.push(
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
                description: "Content List Settings"
            },
            groups: [group, group_Back]
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
                    text: "Required: These fields are required to be filled out. If disabled (grayed-out), \
                        this field is required by the app and cannot be changed."
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

    /** PropertyPane Groups
     * @returns PropertyPane groups array.
     */
    private propertyPaneGroups_OrgChart(): (IPropertyPaneGroup | IPropertyPaneConditionalGroup)[] {
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

        return [group_Colors];
    }

    /** PropertyPane Groups
     * @returns PropertyPane groups array.
     */
    private propertyPaneGroups_Broadcast(): (IPropertyPaneGroup | IPropertyPaneConditionalGroup)[] {
        const group = {
            groupName: "Broadcast Settings",
            groupFields: []
        };

        group.groupFields.push(
            PropertyPaneDropdown("broadcastDays", {
                label: "Select duration of broadcasted bios for",
                options: [
                    {key: 3, text: "3 days"},
                    {key: 7, text: "7 days"},
                    {key: 14, text: "14 days"},
                    {key: 30, text: "30 days"},
                    {key: 45, text: "45 days"},
                    {key: 60, text: "60 days"},
                    {key: 90, text: "90 days"},
                    {key: 180, text: "180 days"},
                    {key: 365, text: "365 days"}
                ]
            })
        );

        return [group];
    }

    // EVENT HANDLERS
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
            LOG(`ERROR! Unable to create '${ newListName }' list.`, er);
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
            LOG("ERROR! Unable to reset list permissions.", er);
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
            LOG("ERROR! Unable to update content managers for all items.", er);
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
                const data = await DataFactory.getAllLists(),
                    names = data.map(i => i.Title);

                names.sort();
                this.propertyPane_.data.optListNames = names.map(name => { return {key: name, text: name}; });

                // add a blank option
                this.propertyPane_.data.optListNames.unshift({key: "", text: "-"});
                
            }
        } catch (er) {
            LOG("ERROR! Unable to get list names.", er);
        }

        return this.propertyPane_.data.optListNames;
    }

    /** Gets and populates library name options 
     * @returns Array of existing library names. Formatted: [{key: string, text: string}, ...];
     */
    private async populateLibraryNameOptions(): Promise<IPropertyPaneDropdownOption[]> {
        try {
            if (this.propertyPane_.data.optLibraryNames.length === 0) {
                const data = await DataFactory.getAllLibraries(this.context.pageContext.web.absoluteUrl);

                this.propertyPane_.data.optLibraryNames = data.map(library => { return {key: library.ServerRelativeUrl, text: library.Title}; });

                // add a blank option
                this.propertyPane_.data.optLibraryNames.unshift({key: "", text: "-"});
                
            }
        } catch (er) {
            LOG("ERROR! Unable to get library names.", er);
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
            LOG("ERROR! Unable to get site groups.", er);
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
            LOG("ERROR! Unable to get Org Chart Colors.", er);
        }

        return this.propertyPane_.data.optOrgChartColors;
    }
//#endregion
}


//#region GLOBAL (NON-REACT) HELPERS -  REACT TYPE, GLOBAL HELPERS ARE LOCATED ON AboutUsApp.tsx
/** Prints out log messages. Decorated console.info() or console.error() method.
 * @param args Message or object to view in the console. If message starts with "ERROR", DEBUG will use console.error().
 */
export function LOG(...args: any[]) {
    // is an error message, if first argument is a string and contains "error" string.
    const isError = (args.length > 0 && (typeof args[0] === "string")) ? args[0].toLowerCase().indexOf("error") > -1 : false;

    if (window && "console" in window) {
        if (isError && console.error) {
            console.error.apply(null, args);

        } else if (console.info) {
            console.info.apply(null, args);

        }
    }
}

/** Prints out DEBUG messages with StackTrace to the console. Decorated console.trace().
 * Use DEBUG() instead of LOG() to make it easier to find and comment out debug statements before production build.
 * @param args Message or object to view in the console
 */
export function DEBUG(...args: any[]) {
    let output = ["DEBUG:"];
    output = output.concat(args);

    if (window && "console" in window) {
        if (console.trace) {
            console.trace.apply(null, output);
            
        } else if (Error) {
            try{
                const error = new Error();
                if (error.stack) console.info("StackTrace: Not an", error.stack);
            } catch (er) { /* fail silently */ }

        } else {
            console.info.apply(null, output);
        }
    }
}

/** Prints out DEBUG messages to the console. Decorated console.info().
 * @param args Message or object to view in the console
 */
 export function DEBUG_NOTRACE(...args: any[]) {
    let output = ["DEBUG:"];
    output = output.concat(args);

    if (window && "console" in window) {
        console.info.apply(null, output);
    }
}


/** Pauses the script for a set amount of time.
* @param milliseconds Amount of milliseconds to sleep.
* @returns Promise
* @example
* await sleep(1000);  // sleep for 1 second then continue
* // or
* sleep(500).then(() => {});  // sleep for half second then run function
*/
export function sleep(milliseconds: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, milliseconds));
 }
 
/** 
 * @param startDate Date broadcast started
 * @param numDaysToBroadcast Number of days to broadcast
 * @returns True if the broadcast date >= today and is < (broadcast date + number of days to broadcast)
 */
export function isInRange_numDays(date: Date, numDaysToBroadcast: number): boolean {
    if (!date) return false;
    if (!numDaysToBroadcast) numDaysToBroadcast = 0;

    const today = new Date(),
        startDate = new Date(date.toISOString()),
        endDate = new Date(date.setDate(date.getDate() + numDaysToBroadcast));

    return (today >= startDate && today < endDate);
}

/** Check to see if any of the values are in the source array 
 * @param source Array to check for values
 * @param values Array of values
 * @param fnCompare Boolean comparison function for complex values
 * @returns True, if any value is in the source
 */
 export function sourceContainsAny(source: any[], values: any[], fnCompare?: (source: any[], value: any)=>boolean): boolean {
    for (let i = 0; i < values.length; i++) {
        const value = values[i];

        if (fnCompare) {
            if (fnCompare(source, value)) return true;
        } else {
            if (source.indexOf(value) > -1) return true;
        }
    }
    return false;
}

/** Properly escape special characters with backslashes (\).
 * @param str String to escape
 * @returns Escaped string
 */
export function escapeProperly(str: string): string {
    return (str) ? str.replace(/([()[{*+.$^\\|?])/g, '\\$1') : str;
}

/** Convert an array of strings into a single sting. Similar to join(", ") but adds a union (" and ", " or ") to the end.
 * @param arr Array to convert into a string list
 * @param joinSeparator Main separateor between words. Default: ", "
 * @param union Union type. Usually ' and ' or ' or '. Default: " and "
 * @param oxford Add the trailing separator (usually a comma) before the union. Default: true
 * @returns String of words
 */
export function arrayListFormat(arr: string[], joinSeparator: string = ", ", union: string = " and ", oxford: boolean = true): string {
    const array = [...arr],
        last = array.pop();

    if (array.length <= 1) return array.join("");

    return array.join(joinSeparator) + ((oxford && array.length > 1) ? trimEnd(joinSeparator): "") + union + last;
}
//#endregion