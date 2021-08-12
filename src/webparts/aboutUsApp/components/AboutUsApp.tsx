import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';

import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import { Form } from '@pnp/sp/forms';

import DataFactory from './DataFactory';
import CustomDialog from './CustomDialog';
import * as FormControls from './FormControls';
import AboutUsForm, { IAboutUsFormProps } from "./AboutUsForm";

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
           2. this.state cannot be called directly after setting it. It requires a an extra moment to update. i.g.: setTimeout(()=>{}, 0);
           3. setState callback (2nd argument), gets called but state may not be updated. 
        */
    }
    //#endregion

    //#region RENDER
    public render(): React.ReactElement<IAboutUsAppProps> {
        return (
            <div className={styles.aboutUsApp}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <span className={styles.title}>List: { this.props.list.title }</span>
                            <p className={styles.subTitle}>Mode: { this.state.displayType }</p>
                            <p className={styles.subTitle}>JCode: { this.state.jcode }</p>
                        </div>
                    </div>
                    { !this.props.list.exists ? this.createConfigureForm() : null }
                    {/* { this.state.displayType === "page" ? this.createPageDisplay() : null } */}
                    {/* { this.state.displayType === "orgchart" ? this.createOrgChartDisplay() : null } */}
                    {/* { this.state.displayType === "accordian" ? this.createAccordianDisplay() : null } */}
                    {/* { this.state.displayType === "phone" ? this.createPhoneDisplay() : null } */}
                    { this.state.displayType === "new" ? <AboutUsForm ctx={ AboutUsApp.ctx } list={ this.props.list } form="new" /> : null }
                    {/* { this.state.displayType === "edit" ? this.createEditForm() : null } */}
                </div>
            </div>
        );

    }

    private createConfigureForm(): React.ReactElement {
        return <FormControls.ShowConfigureWebPart
                onConfigure={ AboutUsApp.ctx.propertyPane.open }
                description="'About-Us' app requires a content list. Select or create one from the settings pane. Click 'Configure' to edit the web part's properties."
            />;
    }
    //#endregion

    //#region 

    //#endregion

    //#region 

    //#endregion
}
