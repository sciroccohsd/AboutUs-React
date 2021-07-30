import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConditionalGroup,
    IPropertyPaneConfiguration,
    IPropertyPaneDropdownOption,
    IPropertyPaneField,
    IPropertyPaneGroup,
    PropertyPaneButton,
    PropertyPaneButtonType,
    PropertyPaneDropdown,
    PropertyPaneDropdownOptionType,
    PropertyPaneLabel,
    PropertyPaneLink,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AboutUsAppWebPartStrings';
import AboutUsApp, { IAboutUsAppProps } from './components/AboutUsApp';

import { SPComponentLoader } from "@microsoft/sp-loader";
import DataFactory, { IListValidationResults } from './components/DataFactory';
import CustomDialog from './components/CustomDialog';
import { trim, escape } from 'lodash';

export interface IAboutUsAppWebPartProps {
    displayType: string;
    displayTypeOptions: Partial<IPropertyPaneDropdownOption[]>;
    listName: string;
    ppListName_dropdown: string | number;
    orgchart_key: { [color: string]: string };
}


enum PROPERTYPANE_STATE {
    "init",
    "update_list",
    "select_list",
    "ready",
    "loading"
}

export default class AboutUsAppWebPart extends BaseClientSideWebPart<IAboutUsAppWebPartProps> {
    //#region PROPERTIES
    private list_: DataFactory;     // Content List object with methods and helpers

    // private listStatus_: STATUS;    // because the list_.exists() is an AJAX request, this variable is used to store the results for use in non-AJAX methods
    private propertyPane_ = {
        "defaultName": "AboutUs_Content",           // suggested list name. may not be the list this app is using.
        "listViewState": PROPERTYPANE_STATE.init,       // what state the property pane view is in for the Content List section
        "listNames": []             // the Web's lists names. stored values from the AJAX request
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

        //SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
    }

    protected async onInit(): Promise<void> {
        return await super.onInit().then(async _ => {
            // create the 'About-Us' list (datafactory)
            this.list_ = new DataFactory(this.context);

            // check if list exists; may have gotten deleted, renamed, or never initialized
            // in the network console, this request may return a 404 status if the list doesn't exist.
            const exists = await DataFactory.listExists(this.properties.listName);
            if (exists) {
                // list exists! now ensure the list properties and fields are still configured
                try {
                    await this.list_.ensureList(this.properties.listName);

                } catch (er) {
                    AboutUsAppWebPart.DEBUG(`ERROR! Could not ensure '${ this.properties.listName }' list properties or fields. ` +
                        `Check to see if the list exists and you have the proper permissions.`);
                }

                await this.list_.test(this.properties.orgchart_key);

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
                list: this.list_
            }
        );

        ReactDom.render(ELEMENT, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    /**
     * Update properties that affect the display.
     * @param properties Key/value pair of property(s) that changed. Properties that affects the display.
     */
    private async updateRenderProperty(propertyName: "listName" | "displayType", value: any, renderNow: boolean = true): Promise<void> {
        
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
    private async getListNames(force: boolean = false): Promise<string[]> {
        // get lists
        if (this.propertyPane_.listNames.length === 0 || force) {
            const DATA = await DataFactory.getAllLists();

            this.propertyPane_.listNames.length = 0;
            this.propertyPane_.listNames = DATA.map(i => i.Title);
            this.propertyPane_.listNames.sort();

        }

        return this.propertyPane_.listNames;
    }

    /**
     * 'Create New List' click event handler. Prompts user for a list name, then creates it.
     * @param evt Click event object
     * @returns Promise. Resolved when the dialog closes and the list is created
     */
    private async createList_click(evt: Event): Promise<void> {
        let newListName: string = this.properties.listName || this.propertyPane_.defaultName;
        let validation: IListValidationResults = { "valid": false, "message": "" };

        // update property pane view
        this.propertyPane_.listViewState = PROPERTYPANE_STATE.loading;
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
                this.propertyPane_.listViewState = PROPERTYPANE_STATE.select_list;
                this.context.propertyPane.refresh();
                return;
            }

            // clean up (trim) use input
            newListName = trim(newListName);
        } while ((validation = DataFactory.validateListName(newListName, this.propertyPane_.listNames)).valid === false && newListName !== null);

        // try to create the list and fields
        const modalMsg = CustomDialog.modalMsg.apply(null, this.modalMsg_processing);
        try {
            await this.list_.ensureList(newListName);
            await this.updateRenderProperty("listName", newListName);
            this.propertyPane_.listViewState = PROPERTYPANE_STATE.update_list;

        } catch (er) {
            AboutUsAppWebPart.DEBUG(`ERROR! Unable to create '${ newListName }' list.`, er);
            await CustomDialog.alert("Something went wrong! See the console for details.", "ERROR");
            this.propertyPane_.listViewState = PROPERTYPANE_STATE.select_list;
            
        }
        modalMsg.close();

        // refresh property pane
        this.context.propertyPane.refresh();

    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        let self = this;
        let propertyPaneGroups: (IPropertyPaneGroup | IPropertyPaneConditionalGroup)[] = [];
        let groupFields_general: IPropertyPaneField<any>[] = [];
        let groupFields_orgchart: IPropertyPaneField<any>[] = [];
        let groupFields_list: IPropertyPaneField<any>[] = [];

        // Display Type:
        groupFields_general.push(
            PropertyPaneDropdown("displayType", {
                label: "Display Type",
                options: this.properties.displayTypeOptions
            })
        );

        // Org Chart: show if display type is OrgChart & content list was ensured
        if (this.properties.displayType === "OrgChart" && this.list_.exists) {
            const colors = this.list_.getOrgChartColors();
            if (colors.length > 0) {
                // label: Key
                groupFields_orgchart.push(
                    PropertyPaneLabel("lblOrgChartColor", {
                        text: "Color Key:"
                    })
                );

                // create each color label & field
                colors.forEach( color => {
                    // the orgchart_key is lowercased & no spaces
                    const colorKey = trim(color).replace(/\s/g, "").toLowerCase();

                    // ensure the color key exists in orgchart_key
                    if (!Object.prototype.hasOwnProperty.call(this.properties.orgchart_key, colorKey)) {
                        this.properties.orgchart_key[colorKey] = "";
                    }
                    // textbox: [COLOR-KEY]
                    groupFields_orgchart.push(
                        PropertyPaneTextField(`orgchart_key.${ colorKey }`, {
                            label: `${ color }:`
                        })
                    );
                });

            }
        }

        // Content List:
        switch (this.propertyPane_.listViewState) {
            case PROPERTYPANE_STATE.init:           // initialize, get data
                this.propertyPane_.listViewState = PROPERTYPANE_STATE.loading;

                // check to see if list exists and get all list names
                this.getListNames().then( responses => {
                    this.propertyPane_.listViewState = (this.list_.exists) ? PROPERTYPANE_STATE.update_list : PROPERTYPANE_STATE.select_list;
                    this.context.propertyPane.refresh();
                });

                break;

            case PROPERTYPANE_STATE.select_list:   // list doesn't exists or selecting/creating new list
                const dropdownOptions: IPropertyPaneDropdownOption[] = this.propertyPane_.listNames.map(x => { return { key: x, text: x }; });
                dropdownOptions.unshift({ key: "", text: "-" });

                // button: Create List
                groupFields_list.push(
                    PropertyPaneButton("btnCreateNewList", {
                        buttonType: PropertyPaneButtonType.Normal,
                        text: "Create New List",
                        onClick: this.createList_click.bind(self)
                    })
                );

                //label: or
                groupFields_list.push(
                    PropertyPaneLabel("lblOr", {
                        text: " or"
                    })
                );

                // dropdown: Select from existing list
                groupFields_list.push(
                    PropertyPaneDropdown("ppListName_dropdown", {
                        label: 'Select from an existing list:',
                        options: dropdownOptions
                    })
                );

                if (this.list_.exists) {
                    // label: back
                    groupFields_list.push(
                        PropertyPaneLabel("lblBackToUpdate", {
                            text: "Back"
                        })
                    );

                    // button: back to update_list
                    groupFields_list.push(
                        PropertyPaneButton("btnBackToUpdate", {
                            buttonType: PropertyPaneButtonType.Icon,
                            icon: "Back",
                            text: "Back",
                            onClick: () => {
                                this.propertyPane_.listViewState = PROPERTYPANE_STATE.update_list;
                                this.context.propertyPane.refresh();
                            }
                        })
                    );
                }

                break;

            case PROPERTYPANE_STATE.update_list:    // list exists

                // label: current list
                groupFields_list.push(
                    PropertyPaneLabel("lblCurrentList", {
                        text: "Current list:"
                    })
                );

                // link: current list
                groupFields_list.push(
                    PropertyPaneLink("lnkListName", {
                        text: this.properties.listName,
                        href: this.list_.properties.RootFolder.ServerRelativeUrl,
                        target: "_blank"
                    })
                );

                // label: Update list
                groupFields_list.push(
                    PropertyPaneLabel("lblUpdateList", {
                        text: "Update list properties & fields:"
                    })
                );

                // button: Update list
                groupFields_list.push(
                    PropertyPaneButton("btnUpdateList", {
                        text: "Update now!",
                        onClick: async () => {
                            this.propertyPane_.listViewState = PROPERTYPANE_STATE.loading;
                            this.context.propertyPane.refresh();

                            const modalMsg = CustomDialog.modalMsg.apply(null, this.modalMsg_processing);
                            try {
                                await this.list_.ensureList(this.list_.title);
                                modalMsg.close();

                            } catch (er) {
                                modalMsg.close();
                                AboutUsAppWebPart.DEBUG(`ERROR! Unable to update '${ this.properties.listName }' list with About-Us properties & fields.`, er);
                                await CustomDialog.alert("Something went wrong! See the console for details.", "ERROR");

                            }

                            this.propertyPane_.listViewState = PROPERTYPANE_STATE.update_list;
                            this.context.propertyPane.refresh();
                        }
                    })
                );

                // label: Select a list
                groupFields_list.push(
                    PropertyPaneLabel("lblSelectList", {
                        text: "Select or create a new list:"
                    })
                );

                // button: Select a list
                groupFields_list.push(
                    PropertyPaneButton("btnSelectList", {
                        text: "Choose a list",
                        onClick: () => {
                            this.propertyPane_.listViewState = PROPERTYPANE_STATE.select_list;
                            this.context.propertyPane.refresh();
                        }
                    })
                );

                break;

        }

        // add propertypane groups
        if (groupFields_general.length > 0) propertyPaneGroups.push({ "groupName": "General", "groupFields": groupFields_general });
        if (groupFields_orgchart.length > 0) propertyPaneGroups.push({ "groupName": "Org Chart", "groupFields": groupFields_orgchart });
        if (groupFields_list.length > 0) propertyPaneGroups.push({ "groupName": "About-Us List", "groupFields": groupFields_list });

        // return property pane configuration
        return {
            showLoadingIndicator: this.propertyPane_.listViewState === PROPERTYPANE_STATE.loading,
            pages: [
                {
                    header: {
                        description: "Settings"
                    },
                    groups: propertyPaneGroups
                }
            ]
        };
    }

    /**
     * PropertyPane onChange event handler. Triggers anytime a PropertyPane form element is changed.
     * @param propertyPath PropertyPane's target name. Use this property to determine which element triggered the onChange.
     * @param oldValue Previous value
     * @param newValue New value
     */
    public onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {

        // do something based on which property (propertyPath) changed
        switch (propertyPath) {
            case "displayType": // display type dropdown changed
                this.updateRenderProperty("displayType", newValue);
                this.context.propertyPane.refresh();
                break;

            case "ppListName_dropdown":  // select a list dropdown changed
                
                if (newValue === "") return;

                this.propertyPane_.listViewState = PROPERTYPANE_STATE.loading;
                this.context.propertyPane.refresh();

                // show warning. we need to ensure the 'About-Us' list properties and fields are set/available
                CustomDialog.confirm(
                    `Click 'Continue' to update '${ newValue }' list with About-Us properties and fields.`,
                    "Update list properties and fields?",
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
                                this.propertyPane_.listViewState = PROPERTYPANE_STATE.update_list;
                                this.context.propertyPane.refresh();
                                modalMsg.close();

                            }).catch( er => {
                                modalMsg.close();
                                AboutUsAppWebPart.DEBUG(`ERROR! Unable to check or update '${ newValue }' list with About-Us properties & fields.`, er);
                                CustomDialog.alert("Something went wrong! See the console for details.", "ERROR").then( () => {
                                    this.properties.ppListName_dropdown = "";
                                    this.propertyPane_.listViewState = PROPERTYPANE_STATE.select_list;
                                    this.context.propertyPane.refresh();
                                });

                            });

                        } else {
                            // cancelled
                            this.properties.ppListName_dropdown = "";
                            this.propertyPane_.listViewState = PROPERTYPANE_STATE.select_list;
                            this.context.propertyPane.refresh();

                        }
                    });

                break;
        
        }
    }
    //#endregion

    //#region HELPERS
    /**
     * Prints our debug messages. Decorated console.info() or console.error() method.
     * @param args Message or object to view in the console. If message starts with "ERROR", DEBUG will use console.error().
     */
    public static DEBUG(...args: any[]) {
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
