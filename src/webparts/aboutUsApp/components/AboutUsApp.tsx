import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';

import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import { Form } from '@pnp/sp/forms';

import DataFactory from './DataFactory';
import * as FormControls from './FormControls';
import AboutUsForm, { IAboutUsFormProps } from "./AboutUsForm";
import { IAboutUsAppWebPartProps } from '../AboutUsAppWebPart';
import AboutUsDisplay from './AboutUsDisplay';

//#region INTERFACES, TYPES & ENUMS
export interface IAboutUsAppProps {
    displayType: string;
    properties: IAboutUsAppWebPartProps;
    list: DataFactory;
}

interface IAboutUsAppState {
    displayType: string;
    itemId: number;
}
//#endregion


export default class AboutUsApp extends React.Component<IAboutUsAppProps, IAboutUsAppState, {}> {
//#region PROPERTIES
    public static ctx: WebPartContext = null;   // must set the context before using this class

    private formValues = {};
//#endregion

//#region RENDER
    constructor(props) {
        super(props);

        const url = new URL(window.location.href);
        let form = (url.searchParams.get("form") || "").toLowerCase(),
            itemId = parseInt((url.searchParams.get("id") || "0"), 10);

        // make sure 'display' param value is valid:
        if (["new", "edit"].indexOf(form) === -1) form = "";

        // initialize state
        this.state = {
            "displayType": form || this.props.displayType || "page",
            "itemId": (itemId > 0) ? itemId : null
        };
        /* Observations on this.setState
           1. Doesn't update objects unless it is reassigned (e.g.: Object.assign({}, this.state.###)).
           2. this.state cannot be called directly after setting it. It requires an extra moment to update. i.g.: setTimeout(()=>{}, 0);
           3. setState callback (2nd argument), gets called but state may not be updated. 
        */
    }

    public render(): React.ReactElement<IAboutUsAppProps> {
        // DEBUG
        LOG("this.props.lists:", this.props.list);

        return (
            <div className={styles.aboutUsApp}>
                { !this.props.list.exists ? this.createConfigureForm() : 
                    <div>
                        { this.state.displayType === "page" ? this.createPageDisplay() : null }
                        {/* { this.state.displayType === "orgchart" ? this.createOrgChartDisplay() : null } */}
                        {/* { this.state.displayType === "accordian" ? this.createAccordianDisplay() : null } */}
                        {/* { this.state.displayType === "phone" ? this.createPhoneDisplay() : null } */}
                        {/* { this.state.displayType === "datatable" ? this.createDatatableDisplay() : null } */}
                        {/* { this.state.displayType === "broadcast" ? this.createBroadcastDisplay() : null } */}
                        { this.state.displayType === "new" ? this.createNewForm() : null }
                        { this.state.displayType === "edit" ? this.createEditForm() : null }
                    </div>
                }
            </div>
        );

    }

    private createConfigureForm(): React.ReactElement {
        return <FormControls.ShowConfigureWebPart
                onConfigure={ AboutUsApp.ctx.propertyPane.open }
                iconText="Configure About-Us Web Part"
                description="'The About-Us app requires a content list. Open the settings pane to create or select a content list."
                buttonLabel="Settings"
            />;
    }

    private createPageDisplay(): React.ReactElement {
        return React.createElement(AboutUsDisplay, {
            ctx: AboutUsApp.ctx,
            properties: this.props.properties,
            list: this.props.list,
            itemId: this.state.itemId,
            history: History
        });
    }

    private createNewForm(): React.ReactElement {
        return React.createElement(AboutUsForm, {
            ctx: AboutUsApp.ctx,
            properties: this.props.properties,
            list: this.props.list,
            form: "new",
            history: History,
        });
    }

    private createEditForm(): React.ReactElement {
        return React.createElement(AboutUsForm, {
            ctx: AboutUsApp.ctx,
            properties: this.props.properties,
            list: this.props.list,
            form: "edit",
            itemId: this.state.itemId,
            history: History,
        });
    }
//#endregion
}


// export function Wrapper({condition, wrapper, children}) {
//     return (condition) ? wrapper(children) : children;
// }
export interface IWrapperProps {
    condition: any;
    wrapper: (children) => any;
    children: any;
    elseWrapper?: (children) => any;
}
export class Wrapper extends React.Component<IWrapperProps> {
    public render(): React.ReactElement<any> {
        return (this.props.condition)
            ? this.props.wrapper(this.props.children)
            : (this.props.elseWrapper)
                ? this.props.elseWrapper(this.props.children)
                : this.props.children ;
    }
}

//#region PRIVATE LOG
/** Prints out debug messages. Decorated console.info() or console.error() method.
 * @param args Message or object to view in the console. If message starts with "ERROR", DEBUG will use console.error().
 */
 function LOG(...args: any[]) {
    // is an error message, if first argument is a string and contains "error" string.
    const isError = (args.length > 0 && (typeof args[0] === "string")) ? args[0].toLowerCase().indexOf("error") > -1 : false;
    args = ["(About-Us AboutUsApp.tsx)"].concat(args);

    if (window && window.console) {
        if (isError && console.error) {
            console.error.apply(null, args);

        } else if (console.info) {
            console.info.apply(null, args);

        }
    }
}
//#endregion