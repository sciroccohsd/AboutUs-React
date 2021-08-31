import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';

import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import { Form } from '@pnp/sp/forms';

import DataFactory from './DataFactory';
import CustomDialog from './CustomDialog';
import * as FormControls from './FormControls';
import AboutUsForm, { IAboutUsFormProps } from "./AboutUsForm";
import AboutUsForms from './AboutUsForm';

export interface IAboutUsAppProps {
    displayType: string;
    list: DataFactory;
}

interface IAboutUsAppState {
    displayType: string;
    jcode: string;
    items: any;
    item?: any;
}

export default class AboutUsApp extends React.Component<IAboutUsAppProps, IAboutUsAppState, {}> {
    //#region PROPERTIES
    public static ctx: WebPartContext = null;   // must set the context before using this class

    private formValues = {};
    //#endregion

    //#region CONSTRUCTOR
    constructor(props) {
        super(props);

        const url = new URL(window.location.href);
        let form = (url.searchParams.get("form") || "").toLowerCase(),
            jcode = (url.hash || url.searchParams.get("jcode") || "").toLowerCase();

        // make sure 'display' param value is valid:
        if (["new", "edit"].indexOf(form) === -1) form = "";

        // initialize state
        this.state = {
            "displayType": form || this.props.displayType,
            "jcode": jcode,
            "items": [],
            "item": null
        };
        /* Observations on this.setState
           1. Doesn't update objects unless it is reassigned (e.g.: Object.assign({}, this.state.###)).
           2. this.state cannot be called directly after setting it. It requires an extra moment to update. i.g.: setTimeout(()=>{}, 0);
           3. setState callback (2nd argument), gets called but state may not be updated. 
        */
    }
    //#endregion

    //#region RENDER
    public render(): React.ReactElement<IAboutUsAppProps> {
        return (
            <div className={styles.aboutUsApp}>
                <div className={styles.container}>
                    { !this.props.list.exists ? this.createConfigureForm() : 
                        <div>
                            {/* { this.state.displayType === "page" ? this.createPageDisplay() : null } */}
                            {/* { this.state.displayType === "orgchart" ? this.createOrgChartDisplay() : null } */}
                            {/* { this.state.displayType === "accordian" ? this.createAccordianDisplay() : null } */}
                            {/* { this.state.displayType === "phone" ? this.createPhoneDisplay() : null } */}
                            { this.state.displayType === "new" ? React.createElement(AboutUsForm, {
                                    ctx: AboutUsApp.ctx,
                                    list: this.props.list,
                                    form: "new",
                                    history: History,
                                }) : null 
                            }
                            {/* { this.state.displayType === "edit" ? this.createEditForm() : null } */}
                        </div>
                    }
                </div>
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
    //#endregion

    //#region 

    //#endregion

    //#region 

    //#endregion
}
