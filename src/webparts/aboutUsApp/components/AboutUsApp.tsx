import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';

import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import { Form } from '@pnp/sp/forms';

import DataFactory from './DataFactory';
import * as FormControls from './FormControls';
import AboutUsForm, { IAboutUsFormProps } from "./AboutUsForm";
import { DEBUG, DEBUG_NOTRACE, IAboutUsAppWebPartProps, LOG } from '../AboutUsAppWebPart';
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
        let form = this.getAboutUsForm(),
            itemId = this.getAboutUsID();

        // make sure 'display' param value is valid:
        if (["new", "edit"].indexOf(form) === -1) form = "";

        // initialize state
        this.state = {
            "displayType": form || this.props.displayType,
            "itemId": itemId
        };

        // set initial history state
        history.replaceState(this.state, document.title, window.location.href);

        // handle browser back button to reduce page refreshes
        window.onpopstate = this.window_onpopstate.bind(this);
    }

    public render(): React.ReactElement<IAboutUsAppProps> {
        DEBUG_NOTRACE("this.state:", this.state);
        DEBUG_NOTRACE("this.props.lists:", this.props.list);

        return (
            <div className={styles.aboutUsApp}>
                { !this.props.list.exists ? this.createConfigureForm() : 
                    <>
                        { this.state.displayType === "page" ? this.createPageDisplay() : null }
                        {/* { this.state.displayType === "orgchart" ? this.createOrgChartDisplay() : null } */}
                        {/* { this.state.displayType === "accordian" ? this.createAccordianDisplay() : null } */}
                        {/* { this.state.displayType === "phone" ? this.createPhoneDisplay() : null } */}
                        {/* { this.state.displayType === "datatable" ? this.createDatatableDisplay() : null } */}
                        {/* { this.state.displayType === "broadcast" ? this.createBroadcastDisplay() : null } */}
                        { this.state.displayType === "new" ? this.createNewForm() : null }
                        { this.state.displayType === "edit" ? this.createEditForm() : null }
                    </>
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
            changeDisplay: this.changeDisplayType.bind(this),
            changeItem: this.changeItemID.bind(this)
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

    private getAboutUsID(): number {
        const url = new URL(window.location.href),
            id = parseInt(url.searchParams.get(this.props.properties.urlParam) || "0", 10);

        return (id > 0) ? id : null;
    }

    private getAboutUsForm(): string {
        const url = new URL(window.location.href),
            displayType = url.searchParams.get(`${this.props.properties.urlParam}form`);

        return (displayType) ? displayType.toLowerCase() : "";
    }
//#endregion


//#region DISPLAY CHANGES
    private changeDisplayType(displayType: string) {
        // don't change view if the display type didn't change
        if (this.state.displayType === displayType) return;

        const url = new URL(location.href),
            formParam = `${this.props.properties.urlParam}form`;

        // update search param if display "new" or "edit" forms
        // don't push state
        if (displayType === "new" || displayType === "edit") {
            url.searchParams.set(formParam, displayType);
            //return location.assign(url.toString());
        } else {
            url.searchParams.delete(formParam);
        }

        this.setState({...this.state, "displayType": displayType}, () => {
            history.pushState(this.state, document.title, url.toString());
        });
    }

    private changeItemID(id: number, title: string = "", url?: string, replaceState?: boolean) {
        // don't change navigation if the ID didn't change
        if (this.state.itemId === id) return;

        if (!url) {
            const href = new URL(location.href);
            href.searchParams.set(this.props.properties.urlParam, id.toString());
            url = href.toString();
        }

        this.setState({...this.state, "itemId": id}, () => {
            if (replaceState === true) {
                history.replaceState(this.state, document.title || title, url);
            } else {
                history.pushState(this.state, document.title || title, url);
            }
        });
    }

    private window_onpopstate(evt) {
        let state: Partial<IAboutUsAppState> = evt.state;

        if (!state) state = {};
        if (!state.displayType) state.displayType = this.getAboutUsForm() || this.props.displayType;
        if (!state.itemId) state.itemId = this.getAboutUsID();

        if (this.state.displayType !== state.displayType || this.state.itemId !== state.itemId) {
            this.setState(state as IAboutUsAppState);
        }
    }
//#endregion
}


//#region GLOBAL HELPERS (REACT TYPES)
export interface IWrapperProps {
    condition: boolean;
    wrapper: (children) => any;
    children: React.ReactNode | React.ReactNodeArray;
    else?: (children) => any;
}
export class Wrapper extends React.Component<IWrapperProps> {
    public render(): React.ReactElement<any> {
        return (this.props.condition)
            ? this.props.wrapper(this.props.children)
            : (this.props.else)
                ? this.props.else(this.props.children)
                : this.props.children ;
    }
}
//#endregion