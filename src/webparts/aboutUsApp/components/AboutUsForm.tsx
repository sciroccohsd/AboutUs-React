import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';

import DataFactory, { IDataFactoryFieldInfo } from './DataFactory';
import CustomDialog from './CustomDialog';

import * as FormControls from './FormControls';
import { IFormControlDropdownProps, IRichTextControlProps } from './FormControls';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Form } from '@pnp/sp/forms';
//import { IFieldInfo } from '@pnp/sp/fields';
import { IDropdown, IDropdownOption, IDropdownProps, IInputProps, ILabelProps, ITextFieldProps, Stack } from 'office-ui-fabric-react';
import * as strings from 'AboutUsAppWebPartStrings';

export interface IAboutUsFormProps {
    ctx: WebPartContext;
    list: DataFactory;
    form: "new" | "edit";
    jcode?: string;
}

export interface IAboutUsValueState {
    defaultValue: any;  // starting or default value
    value: any;         // current value on form
    errorText: string;  // error message if any, else null or "";
    disabled: boolean;  // field is disabled
}

interface IAboutUsFormState {
    valueStates?: {[key:string]: IAboutUsValueState};
}

export default class AboutUsForms extends React.Component<IAboutUsFormProps, IAboutUsFormState, {}> {
    private refForm = {};
    //#region PROPERTIES

    //#endregion

    //#region CONSTRUCTOR
    constructor(props) {
        super(props);

        this.state = {
            valueStates: null,
        };
    }
    //#endregion
    
    //#region RENDER
    public render(): React.ReactElement<IAboutUsFormProps> {
        return (
            <div className={styles.form}>
                { this.props.form === "new" ? this.newForm() : null }
                { this.props.form === "edit" ? this.editForm(this.props.jcode) : null }
            </div>            
        );
    }
    //#endregion
    
    //#region NEW FORM
    private newForm(): React.ReactElement {
        let elements: React.ReactElement[] = [],
            valueStates: {[key:string]: IAboutUsValueState} = {};

        this.props.list.fields.forEach( field => {
            try{
                let valueState = this.getValueState_by_InternalName(field.InternalName);
                if (valueState === null) {
                    // if null, create the value state
                    valueState = this.initializeNewFieldValueState(field.DefaultValue || "");
                    // keep track of all the newly created value states
                    valueStates[field.InternalName] = valueState;
                }

                ///TODO: custom form controls
                
                // default form controls
                switch (field["odata.type"]) {
                    case "SP.FieldText":
                        elements.push( this.spFieldText(field, valueState) );
                        break;

                    case "SP.FieldMultiLineText":
                        if (field.RichText === true) {
                            // is multiline rich text  field
                            elements.push( this.spRichText(field, valueState) );
                        } else {
                            //  basic multiline field
                            elements.push( this.spFieldText(field, valueState) );
                        }
                        break;
                
                    case "SP.FieldChoice":
                        elements.push( this.spDropDown(field, valueState) );
                        break;
                            
                    case "SP.FieldMultiChoice":
                        valueState.defaultValue = this.parseSPDelimitedStringValues( valueState.defaultValue);
                        valueState.value = this.parseSPDelimitedStringValues( valueState.value );
                        elements.push( this.spMultiDropDown(field, valueState) );
                        break;
                        
                        default:

                        break;
                }

            } catch(er) {
                AboutUsForms.DEBUG("ERROR: newForm()", er);
            }
        });

        // re-initialize fieldValueState state. clear this state when form closes
        if (!this.state.valueStates) this.state = { "valueStates": valueStates };

        return (
            <Stack tokens={{ childrenGap: 10 }}>{ elements }</Stack>
        );
    }

    /**
     * Initialize a new ValueState: {defaultValue, value, errorText}
     * @param id Field InternalName
     * @param value Initial or default value
     * @returns ValueState object
     */
    private initializeNewFieldValueState(value?, disabled: boolean = false): IAboutUsValueState {
        if (value === undefined) value = null;
        return {
            defaultValue: value,
            value: value,
            errorText: null,
            disabled: disabled
        };
    }
    //#endregion
    
    //#region EDIT FORM
    private editForm(jcode: string): React.ReactElement {
        return (
            <div></div>
        );
    }
    //#endregion

    //#region CONTROLS
    /**
     * Create a textbox control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldText(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const isMultiline = field["odata.type"] === "SP.FieldMultiLineText",
            props: ITextFieldProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required,
            label: field.Title,
            placeholder: `Enter ${ field.Title }...`,
            description: field.Description,

            multiline: isMultiline,
            cols: isMultiline ? field.NumberOfLines : null,
            defaultValue: valueState.value,
            maxLength: isMultiline ? null : Math.min((field.MaxLength || 255), 255),

            errorMessage: valueState.errorText,

            onChange: this.input_change.bind(this)
        };

        return React.createElement(FormControls.TextboxControl, props);
    }

    private spRichText(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: IRichTextControlProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required,
            label: field.Title,
            description: field.Description,
            placeholder: `Enter ${ field.Title }...`, // unable to change RichText placeholder styling

            value: valueState.value,

            errorMessage: valueState.errorText,

            onChange: newText => this.richtext_change(newText, field.InternalName)
        };

        return React.createElement(FormControls.RichTextControl, props);
    }

    /**
     * Create a dropdown control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
     private spDropDown(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: IFormControlDropdownProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required || null,
            label: field.Title,
            placeholder: "Select an option",
            description: field.Description,

            options: this.generateDropDownOptions(field.Choices),
            defaultSelectedKey: valueState.value,

            errorMessage: valueState.errorText,

            onChange: this.dropdown_change.bind(this)
        };

        return React.createElement(FormControls.DropdownControl, props);
    }

        /**
     * Create a multi-select dropdown control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spMultiDropDown(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: IFormControlDropdownProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required || null,
            label: field.Title,
            placeholder: "Select an option",
            description: field.Description,

            multiSelect: true,
            options: this.generateDropDownOptions(field.Choices),
            defaultSelectedKeys: valueState.value,

            errorMessage: valueState.errorText,

            onChange: this.dropdown_change.bind(this)
        };

        return React.createElement(FormControls.DropdownControl, props);
    }
    //#endregion
    
    //#region EVENT HANDLERS
    /**
     * INPUT Element onChange() event handler
     * @param evt Event object
     */
    private input_change(evt: React.ChangeEvent<HTMLInputElement>) {
        try {
            const internalName = evt.target.id,
                value = evt.target.value,
                valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            valueState.value = value;

            // validate / set errorText
            valueState.errorText = this.validateFieldValue(internalName, value);

            AboutUsForms.DEBUG("inputElement_change(evt):", internalName, value, this.state.valueStates[internalName]);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! input_change():", evt.target.id, er);
        }
    }

    /**
     * RichText Control event handler.
     * - HINT: Place this event handler inside an arrow function. See example.
     * @param newText HTML string value returned from SP.PnP RichText control
     * @param internalName Field InternalName used as the field's reference ID.
     * @returns HTML string value
     * @example
     * const props = {
     *     onChange: newText => this.richtext_change(newText, field.InternalName)
     * }
     */
    private richtext_change(newText: string, internalName: string): string {
        AboutUsForms.DEBUG("richtext_change(value)", internalName, newText);
        try {
            const valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            valueState.value = newText;

            // validate / set errorText
            valueState.errorText = this.validateFieldValue(internalName, newText);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! richtext_change():", internalName, er);
        }
        
        return newText;
    }
    /**
     * SP Dropdown onChange() event handler
     * @param evt Event object
     * @param value Dropdown value that changed. {key: string, text: string, selected: boolean}
     * @param index Option index that triggered the change
     */
    private async dropdown_change(evt: React.ChangeEvent, value: IDropdownOption, index: number) {
        try{
            const InternalName = evt.target.id,
                field = this.getField_by_InternalName(InternalName),
                valueState = this.getValueState_by_InternalName(InternalName);

            if (valueState === null) return;
            
            if (field["odata.type"] === "SP.FieldMultiChoice") {
                // is multi-choice: value returned is what changed, not what's selected
                // need to update the existing value with what changed
                if (value.selected) {
                    // new value added
                    valueState.value.push(value.key);
                } else {
                    // remove selected value
                    const ndx = valueState.value.indexOf(value.key);
                    if (ndx > -1) valueState.value.splice(ndx, 1);
                }

            } else {
                // is single-choice: value returned is the new value
                valueState.value = value.key;
            }

            // validate / set errorText
            valueState.errorText = this.validateFieldValue(InternalName, valueState.value);

            // update state
            await this.setValueState_for(InternalName, valueState);
            
            AboutUsForms.DEBUG("inputElement_change(evt):", InternalName, value, this.state.valueStates[InternalName]);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! dropdown_change():", evt.target.id, er);
        }
    }
    //#endregion
    
    //#region VALIDATION
    private validateFieldValue(InternalName: string, value): string {
        const field = this.getField_by_InternalName(InternalName),
            errorText = [];

        // validate value based on field

        return errorText.join("\n");
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
        args = ["(About-Us AboutUsForm.tsx)"].concat(args);

        if (window && window.console) {
            if (isError && console.error) {
                console.error.apply(null, args);

            } else if (console.info) {
                console.info.apply(null, args);

            }
        }
    }

    /**
     * Generates the IDropdownOption's array from array of string values.
     * @param choices Array of string values.
     * @returns [{"key": string, "text": string}, ...]
     */
    private generateDropDownOptions(choices: string[]): IDropdownOption[] {
        return choices.map( choice => {
            return {
                key: choice,
                text: choice
            };
        });
    }

    /**
     * Parses SP's (";#") delimited string into an array of values
     * @param delimitedString SP's (";#") delimited string value.
     * @returns Values as an array of strings.
     */
    private parseSPDelimitedStringValues(delimitedString: string | string[]): string[] {
        // if null or empty
        if (typeof delimitedString === "undefined" || delimitedString === null) return [];

        // if SP string using ";#" as the separator. remove leading and trailing separators.
        if (typeof delimitedString === "string" && delimitedString.indexOf(";#") > -1) return trim(delimitedString).replace(/^(;#)|(;#)$/g,"").split(";#");

        // if just a single string value
        if (typeof delimitedString === "string") return [trim(delimitedString)];

        // else selected key is an array. no change neccessary
        return delimitedString;
    }

    /**
     * Get Field information
     * @param id Field InternalName of field to retrieve
     * @returns Field info
     */
    private getField_by_InternalName(InternalName: string): IDataFactoryFieldInfo {
        return find(this.props.list.fields, ["InternalName", InternalName]) || null;
    }

    /**
     * Get the ValueState ({defaultValue, value, errorText}) for this field
     * @param InternalName Field InternalName for the ValueState object to retrieve
     * @returns ValueState object for Field ID passed
     */
    public getValueState_by_InternalName(InternalName: string): IAboutUsValueState {
        return (this.state.valueStates && Object.prototype.hasOwnProperty.call(this.state.valueStates, InternalName)) 
            ? this.state.valueStates[InternalName] 
            : null ;
    }

    /**
     * Update the ValueState for a specific field. Updates the state and display. Pauses, to allow the state to update correctly.
     * @param InternalName Field InternalName for the ValueState object to set.
     * @param valueState New ValueState to set.
     */
    public async setValueState_for(InternalName: string, valueState: IAboutUsValueState): Promise<void> {
        this.setState({ "valueStates": {...this.state.valueStates, InternalName: valueState} });
        // setState needs just a momemt to update the state. 
        // not required but helps if you need to reference it right away.
        await this.sleep(0);
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
    //#endregion
}