//> npm install @pnp/spfx-controls-react --save --save-exact
import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { trim, escape } from 'lodash';

import * as ReactControls from "@pnp/spfx-controls-react";

import DataFactory from './DataFactory';
import CustomDialog from './CustomDialog';
import { Field, IFieldInfo } from '@pnp/sp/fields';

export interface IFormControlsProps {
    displayType: "display" | "edit" | "new";
    controlType?: string;    // custom control name

    field: IFieldInfo;
    value?: any;
    className?: string | string[];
    cssProps?: React.CSSProperties;
}

export default class FormControls extends React.Component<IFormControlsProps, {}> {
    //#region PROPERTIES
    private value: any = null;
    private className: string = "";
    private cssProps: React.CSSProperties;
    //#endregion

    //#region CONSTRUCTOR
    constructor(props) {
        super(props);

        this.state = {};
    }
    //#endregion

    //#region RENDER
    public render(): React.ReactElement<IFormControlsProps> {
        return this.createControl();
    }
    
    /**
     * Create the proper control based on the display & field type.
     * This function acts like a router, it doesn't generate any elements.
     */
    private createControl(): React.ReactElement {
        let elem: React.ReactElement = null;

        // normalize properties
        this.value = this.props.value;
        this.className = FormControls.convertStringStringArray(this.props.className, " ");
        this.cssProps = this.props.cssProps;

        // route
        switch (this.props.displayType) {
            // case "new":
            //     // set value or field's DefaultValue as the default value
            //     if ((this.value === undefined || this.value === null) && 
            //         (this.props.field.DefaultValue !== undefined || this.props.field.DefaultValue !== null)) {
            //             this.value = this.props.field.DefaultValue;
            //     }
            case "edit":
                
                
                break;
        
            default:    // display

                elem = this.textField(this.props.field, "test value", "testClass", { "backgroundColor": "transparent"});
                break;
        }

        return elem;
    }
    //#endregion

    //#region DISPLAY COMPONENTS
    private textField(field: IFieldInfo, value: string = "", className: string = "", cssProps: React.CSSProperties = null): React.ReactElement {
        // normalize arguments

        return <ReactControls.FieldTextRenderer text={ value } className={ className } cssProps={ cssProps } />;
    }
    //#endregion

    //#region EDIT COMPONENTS
    //#endregion

    //#region HELPERS
    /**
     * Convert a 'string | string[]' type to a string.
     * @param strArray String or String Array to convert to a string.
     * @param join String to use as the joiner. Default: "";
     * @returns String
     */
    private static convertStringStringArray(strArray: string | string[], join: string = ""): string {
        // if string array
        if (strArray instanceof Array) return strArray.join(join);
        
        // else 
        return strArray;
    }
    //#endregion
}
