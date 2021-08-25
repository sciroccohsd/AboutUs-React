import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';
import * as strings from 'AboutUsAppWebPartStrings';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import DataFactory, { IDataFactoryFieldInfo, IFieldUrlValue, IUserInfoItem } from './DataFactory';
import CustomDialog from './CustomDialog';
import * as FormControls from './FormControls';

import { IDropdownOption,
    //IPersonaProps,
    ITextFieldProps, 
    Stack } from 'office-ui-fabric-react';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
//import { IPersonaProps } from "@pnp/spfx-controls-react/node_modules/office-ui-fabric-react/lib/components/Persona/Persona.types";
import { BaseComponentContext } from '@pnp/spfx-controls-react/node_modules/@microsoft/sp-component-base';
import { DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react';

export interface IAboutUsFormProps {
    ctx: WebPartContext;
    list: DataFactory;
    form: "new" | "edit";
    jcode?: string;
}

// Form field's current value, default/initial value, error status, disabled status
export interface IAboutUsValueState {
    defaultValue: any;  // starting or default value
    value: any;         // current value on form
    errorMessage: string;  // error message if any, else null or "";
    disabled: boolean;  // field is disabled
}

export interface IAboutUsLookupItemValue {
    sp: {results: any[]} | any;
    control: any[];
}

export interface IAboutUsUserValue {
    sp: { results: number[] } | number;  // user.Id(s)
    control: string[];   // [user.Login || user.Email, ...]
}

enum DISPLAY_STATE {
    "loading",
    "ready"
}

interface IAboutUsFormState {
    valueStates?: {[key:string]: IAboutUsValueState};
    display: DISPLAY_STATE;
}


export default class AboutUsForms extends React.Component<IAboutUsFormProps, IAboutUsFormState, {}> {
//#region PROPERTIES
    private baseComponentContext: BaseComponentContext = null;

    private _div = document.createElement("div");
//#endregion

//#region CONSTRUCTOR
    constructor(props) {
        super(props);

        this.state = {
            valueStates: null,
            display: DISPLAY_STATE.loading
        };
    }
//#endregion
    
//#region RENDER
    public async componentDidMount() {
        if (this.props.form === "new") {
            this.setState({"display": DISPLAY_STATE.ready});
            return;
        } else {
            // edit form: need to get data first
        }
    }
    public render(): React.ReactElement<IAboutUsFormProps> {
        this.baseComponentContext = this.props.ctx as any;
        return (
            <div className={styles.form}>
                { this.state.display === DISPLAY_STATE.ready ? 
                    <form>
                        { this.props.form === "new" ? this.newForm() : null }
                        { this.props.form === "edit" ? this.editForm(this.props.jcode) : null }
                    </form>
                : 
                    <FormControls.LoadingSpinner/>
                }
            </div>            
        );
    }
//#endregion
    
//#region NEW FORM
    /** Create all the form field controls.
     * @returns Stack of field controls
     */
    private newForm(): React.ReactElement {
        let elements: React.ReactElement[] = [],
            valueStates: {[key:string]: IAboutUsValueState} = {};

        this.props.list.fields.forEach( field => {
            try{
                // init values should be the default values
                let valueState = this.getValueState_by_InternalName(field.InternalName);
                if (valueState === null) {
                    // if null, create the value state
                    valueState = this.initializeFieldValueState(field.DefaultValue || "");
                    // keep track of all the newly created value states
                    valueStates[field.InternalName] = valueState;
                }

                let element = this.createFieldControl(field, valueState);
                if (element) elements.push(element);

            } catch(er) {
                AboutUsForms.DEBUG("ERROR: newForm()", field, er);
            }
        });

        // re-initialize fieldValueState state. clear this state when form closes
        if (!this.state.valueStates) this.setState({"valueStates": valueStates});

        return (
            <Stack tokens={{ childrenGap: 10 }}>{ elements }</Stack>
        );
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
    /** Create field control based on the type of SP field.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private createFieldControl(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        let element: React.ReactElement;

        ///TODO: custom form controls
        
        // default form controls
        switch (field["odata.type"]) {
            case "SP.FieldText":
                element = this.spFieldText(field, valueState);
                break;

            case "SP.FieldNumber":
                valueState.value = this.stringToNumber(valueState.value);
                element = this.spFieldNumber(field, valueState);
                break;

            case "SP.FieldMultiLineText":
                if (field.RichText === true) {
                    // is multiline rich text  field
                    element = this.spRichText(field, valueState);
                } else {
                    //  basic multiline field
                    element = this.spFieldText(field, valueState);
                }
                break;
        
            case "SP.FieldChoice":
                element = this.spFieldChoice(field, valueState);
                break;
                    
            case "SP.FieldMultiChoice":
                valueState.value = this.parseSPDelimitedStringValues( valueState.value );
                element = this.spFieldMultiChoice(field, valueState);
                break;

            case "SP.FieldLookup":
                // valueState.value = selectedId || { results: [selectedId, ...] }
                valueState.value = this.createValueState_LookupItem(valueState.value,  field.InternalName);
                element = this.spFieldLookup(field, valueState);
                break;

            case "SP.FieldUrl":
                /* 
                valueState.value = { 
                    _metadata: {type: "SP.FieldUrlValue"},
                    Url: string,
                    Description: string 
                }
                */
                element = this.spFieldUrl(field, valueState);
                break;

            case "SP.FieldDateTime":
                // valueState.value = string ISO8601 date/time
                element = this.spFieldDateTime(field, valueState);
                break;

            case "SP.FieldUser":
                /*
                valueState.value = {
                    sp: { results: number[] } | number;  // user.Id(s)
                    control: string[];   // [user.Login || user.Email, ...]
                }
                */
                valueState.value = this.createValueState_UserValue(valueState.value, field.InternalName);
                element = this.spFieldUser(field, valueState);
                break;

            default:
                AboutUsForms.DEBUG("Unhandled field control:", field["odata.type"], field.InternalName, field);
                break;
        }

        return element;
    }

    /** Create a textbox control.
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
            rows: isMultiline ? field.NumberOfLines : null,
            defaultValue: valueState.value,
            maxLength: isMultiline ? null : Math.min((field.MaxLength || 255), 255),

            errorMessage: valueState.errorMessage,

            onChange: this.input_onChange.bind(this)
        };

        return React.createElement(FormControls.TextboxControl, props);
    }

    /** Create a Rich Text Editor (RTE) field. Uses (@PnP/SPFX-Controls-React RichText)
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spRichText(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: FormControls.IRichTextControlProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required,
            label: field.Title,
            description: field.Description,
            placeholder: `Enter ${ field.Title }...`, // unable to change RichText placeholder styling

            value: valueState.value,

            errorMessage: valueState.errorMessage,

            onChange: newText => this.richtext_onChange(newText, field.InternalName)
        };

        return React.createElement(FormControls.RichTextControl, props);
    }

    /** Create a number control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldNumber(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: ITextFieldProps = {
            type: "number",
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required,
            label: field.Title,
            placeholder: `Enter ${ field.Title }...`,
            description: field.Description,

            defaultValue: valueState.value,
            min: (typeof field.MinimumValue === "number") ? field.MinimumValue : null,
            max: (typeof field.MaximumValue === "number") ? field.MaximumValue : null,

            errorMessage: valueState.errorMessage,

            onChange: this.input_onChange.bind(this)
        };

        return React.createElement(FormControls.TextboxControl, props);
    }

    /** Create a date/time control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldDateTime(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: FormControls.IDateTimeControlProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required,
            label: field.Title,
            placeholder: `Enter ${field.Title}...`,
            description: field.Description,

            value: valueState.value ? new Date(valueState.value) : null,
            dateConvention: field.DateTimeCalendarType === 0 ? DateConvention.DateTime : DateConvention.Date,
            isMonthPickerVisible: false,
            timeConvention: TimeConvention.Hours24,
            timeDisplayControlType: TimeDisplayControlType.Dropdown,
            showGoToToday: false,
            showLabels: false,

            errorMessage: valueState.errorMessage,
            //onGetErrorMessage: (newDate) => {AboutUsForms.DEBUG("spFieldDateTime > onGetErrorMessage", newDate); return "";},

            onChange: newDate => this.datetime_onChange(newDate, field.InternalName)
        };

        return React.createElement(FormControls.DateTimeControl, props);
    }

    /** Create a dropdown control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
     private spFieldChoice(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: FormControls.IDropdownControlProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required || null,
            label: field.Title,
            placeholder: "Select an option",
            description: field.Description,

            options: this.generateDropDownOptions(field.Choices),
            defaultSelectedKey: valueState.value,

            errorMessage: valueState.errorMessage,

            onChange: this.dropdown_onChange.bind(this)
        };

        return React.createElement(FormControls.DropdownControl, props);
    }

    /** Create a multi-select dropdown control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldMultiChoice(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: FormControls.IDropdownControlProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required || null,
            label: field.Title,
            placeholder: "Select an option",
            description: field.Description,

            multiSelect: true,
            options: this.generateDropDownOptions(field.Choices),
            defaultSelectedKeys: valueState.value,

            errorMessage: valueState.errorMessage,

            onChange: this.dropdown_onChange.bind(this)
        };

        return React.createElement(FormControls.DropdownControl, props);
    }

    /** Create a Lookup control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldLookup(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        /* NOTE: MUST USE LISTITEMPICKER SINCE COMBOBOXLISTITEMPICKER CANNOT ACCEPT DEFAULT VALUE */
        const props: FormControls.IListItemPickerControlProps = {
        // const props: FormControls.IComboBoxListItemPickerControlProps = {
            disabled: valueState.disabled,
            required: field.Required,
            label: field.Title,
            description: field.Description,
            placeholder: "Type to search...",

            // ListItemPickerControl properties
            context: this.baseComponentContext,
            itemLimit: field.AllowMultipleValues ? 99 : 1,

            // ComboBoxListItemPickerControl properties
            // spHttpClient: this.props.ctx.spHttpClient as any,
            // multiSelect:  field.AllowMultipleValues,

            // shared properties
            webUrl: this.props.ctx.pageContext.web.absoluteUrl,
            listId: field.LookupList,
            columnInternalName: field.LookupField,
            defaultSelectedItems: valueState.value.control,

            errorMessage: valueState.errorMessage,

            onSelectedItem: items => this.lookup_onSelectedItem(items, field.InternalName)
        };

        return React.createElement(FormControls.ListItemPickerControl, props);
        // return React.createElement(FormControls.ComboBoxListItemPickerControl, props);
    }

    /** Create a URL + Text control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldUrl(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: FormControls.IUrlControlProps = {
            disabled: valueState.disabled,
            id: field.InternalName,
            required: field.Required,
            label: field.Title,
            description: field.Description,

            defaultValue: valueState.value,

            errorMessage: valueState.errorMessage,

            onChange: this.urlfield_onChange.bind(this)
            
        };
        return React.createElement(FormControls.UrlControl, props);
    }

    /** Create a User control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */    
    private spFieldUser(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const props: FormControls.IPeoplePickerControlProps = {
            context: this.baseComponentContext,
            webAbsoluteUrl: this.props.ctx.pageContext.web.absoluteUrl,
            disabled: valueState.disabled,
            required: field.Required,
            label: field.Title,
            placeholder: `Select ${ field.Title }...`,
            description: field.Description,

            defaultSelectedUsers: valueState.value.control,
            ensureUser: true,   // if false, user.id will be a string & not the user ID number
            personSelectionLimit: field.AllowMultipleValues ? 99 : 1,
            // groupId: field.SelectionGroup,

            errorMessage: valueState.errorMessage,

            onChange: users => this.peoplePicker_onChange((users as IPersonaProps[]), field.InternalName)
        };

        return React.createElement(FormControls.PeoplePickerControl, props);
    }
//#endregion
    
//#region EVENT HANDLERS
    /** Textbox, Textarea, & Number Control onChange() event handler
     * @param evt Event object
     */
    private input_onChange(evt: React.ChangeEvent<HTMLInputElement>) {
        try {
            const internalName = evt.target.id,
                field = this.getField_by_InternalName(internalName),
                value = field.TypeAsString === "Number" ? this.stringToNumber(evt.target.value) : evt.target.value,
                valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;

            valueState.value = value;

            // validate / set errorMessage
            const errorMessage = this.validateFieldValue(internalName, value);

            if (errorMessage !== valueState.errorMessage) {
                valueState.errorMessage = errorMessage;
                this.setValueState_for(internalName, valueState);
            } 

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! input_change():", evt.target.id, er);
        }
    }

    /** RichText Control event handler.
     * - HINT: Place this event handler inside an arrow function. See example.
     * @param value HTML string value returned from SP.PnP RichText control
     * @param internalName Field InternalName used as the field's reference ID.
     * @returns HTML string value
     * @example
     * const props = {
     *     onChange: newText => this.richtext_change(newText, field.InternalName)
     * }
     */
    private richtext_onChange(value: string, internalName: string): string {
        try {
            const valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            valueState.value = value;

            // validate / set errorMessage
            let errorMessage = this.validateFieldValue(internalName, value);

            if (valueState.errorMessage !== errorMessage) {
                valueState.errorMessage = errorMessage;
                this.setValueState_for(internalName, valueState);
            }

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! richtext_change():", internalName, er);
        }
        
        return value;
    }

    /** Dropdown (SP's Choice & MultiChoice) Control onChange() event handler
     * @param evt Event object
     * @param value Dropdown value that changed. {key: string, text: string, selected: boolean}
     * @param index Option index that triggered the change
     */
    private dropdown_onChange(evt: React.ChangeEvent, value: IDropdownOption, index: number) {
        try{
            const internalName = evt.target.id,
                field = this.getField_by_InternalName(internalName),
                valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            
            if (field["odata.type"] === "SP.FieldMultiChoice") {
                if (!(valueState.value instanceof Array)) valueState.value = [];

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

            // validate / set errorMessage
            let errorMessage = this.validateFieldValue(internalName, valueState.value);

            if (valueState.errorMessage !== errorMessage) {
                valueState.errorMessage = errorMessage;
                this.setValueState_for(internalName, valueState);
            }

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! dropdown_change():", evt.target.id, er);
        }
    }

    /** Lookup Control event handler.
     * - HINT: Place this event handler inside an arrow function. See example.
     * @param items Array of selected keys. Keys = lookupItem.Id
     * @param internalName Field InternalName used as the field's reference ID.
     * @example
     * const props = {
     *     onSelectedItem: items => this.lookup_onSelectedItem(items, field.InternalName)
     * }
     */
    private lookup_onSelectedItem(items: any[], internalName: string) {
        try {
            const valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            const value = this.createValueState_LookupItem(items, internalName);

            valueState.value = value;

            // validate / set errorMessage
            let errorMessage = this.validateFieldValue(internalName, value);

            if (valueState.errorMessage !== errorMessage) {
                valueState.errorMessage = errorMessage;
                this.setValueState_for(internalName, valueState);
            }

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! lookup_onSelectedItem():", internalName, er);
        }
    }

    /** URL Control event handler.
     * @param val URL or Text value that changed.
     * @param type "url" or "text", which type of value changed.
     * @param internalName Field InternalName used as the field's reference ID.
     */
    private urlfield_onChange(val: string, type: "url" | "text", internalName: string) {
        try {
            const valueState = this.getValueState_by_InternalName(internalName);
            if (valueState === null) return;

            if (!valueState.value) valueState.value = { __metadata: {type: "SP.FieldUrlValue"}, Url: "", Description: "" };
            
            let value: IFieldUrlValue;
            if (type === "url") {
                value = {...valueState.value, Url: val};
            } else {
                value = {...valueState.value, Description: val};
            }
            valueState.value = value;

            // validate / set errorMessage
            let errorMessage = this.validateFieldValue(internalName, value);

            if (valueState.errorMessage !== errorMessage) {
                valueState.errorMessage = errorMessage;
                this.setValueState_for(internalName, valueState);
            }

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! urlfield_onChange():", internalName, er);
        }
            
    }

    /** DateTime Control event handler.
     * @param newDate Date (object) selected.
     * @param type "url" or "text", which type of value changed.
     * @param internalName Field InternalName used as the field's reference ID.
     * @example
     * const props = {
     *     onChange: newDate => this.datetime_onChange(newDate, field.InternalName)
     * }     
     */
    private datetime_onChange(newDate: Date, internalName: string) {
        try {
            const valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            valueState.value = newDate.toISOString();

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, newDate);

            // must setState for date fields, if not, the date value disappears from the UI
            this.setValueState_for(internalName, valueState);

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! datetime_onChange():", internalName, er);
        }
    }

    /** PeoplePicker Control event handler.
     * @param selectedUsers Array of selected User objects.
     * @param internalName Field InternalName used as the field's reference ID.
     * @example
     * const props = {
     *     onChange: users => this.peoplePicker_onChange(users, field.InternalName)
     * }     
     */    
    private peoplePicker_onChange(users: IPersonaProps[], internalName: string) {
        try {
            const valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;
            const value = this.createValueState_UserValue(users, internalName);

            valueState.value = value;

            // validate / set errorMessage
            let errorMessage = this.validateFieldValue(internalName, value);

            if (valueState.errorMessage !== errorMessage) {
                valueState.errorMessage = errorMessage;
                this.setValueState_for(internalName, valueState);
            }

            AboutUsForms.DEBUG("onChange():", internalName, valueState);

        } catch(er) {
            AboutUsForms.DEBUG("ERROR! peoplePicker_onChange():", internalName, er);
        }
    }
//#endregion
    
//#region VALIDATION
    /** Validates user input
     * @param internalName Field InternalName used as the field's reference ID.
     * @param value Formatted value from control
     * @returns 
     */
    private validateFieldValue(internalName: string, value: any): string {
        const field = this.getField_by_InternalName(internalName),
            valueState = this.getValueState_by_InternalName(internalName);

        let error: string = null;

        switch (field["odata.type"]) {
            case "SP.FieldText":
                error = this.validateString(value, field.Required, field.MaxLength || 255);
                break;

            case "SP.FieldNumber":
                error = this.validateNumber(value, field.Required, field.MinimumValue || null, field.MaximumValue || null);
                break;

            case "SP.FieldMultiLineText":
                error =  this.validateString(this.getTextFromHTMLString(value), field.Required);
                break;
        
            case "SP.FieldChoice":
                error = this.validateChoice(value, field.Required);
                break;
                    
            case "SP.FieldMultiChoice":
                error = this.validateChoice(value, field.Required);
                break;

            case "SP.FieldLookup":
                error  = this.validateChoice(value.control, field.Required);
                break;

            case "SP.FieldUrl":
                error = this.validateString(value.Url, field.Required, 255);
                break;

            case "SP.FieldDateTime":
                error = this.validateDate(value, field.Required);
                break;

            case "SP.FieldUser":
                error = this.validateChoice(value.control, field.Required);
                break;

            default:
                AboutUsForms.DEBUG("Unhandled field validation:", field["odata.type"], field.InternalName, value);
                break;
        }

        // validate value based on field

        return error;
    }
//#endregion
    
//#region HELPERS
    /** User text validation
     * @param value String value to validate
     * @param isRequired Boolean if the field is required/mandatory
     * @param max Maximum length of the string
     * @returns Error message or null
     */
    private validateString(value: string, isRequired: boolean = false, max?: number): string {
        const length = trim(value || "").length;

        if (isRequired && length === 0) return "Required.";

        if (typeof max === "number" && length > max) return `Too long. Maximum length is ${max}.`;

        return null;
    }

    /** User number validation
     * @param value Number value to validate
     * @param isRequired Boolean if the field is required/mandatory
     * @param min Minimum number value
     * @param max Maximum number value
     * @returns Error message or null
     */
    private validateNumber(value: number, isRequired: boolean = false, min?: number, max?: number): string {
        const length = typeof value === "number" ? value.toString() : 0;

        if (isRequired && length === 0) return "Required.";

        if (value === null) return null;

        if (typeof min === "number" && value < min) return `Too low. Minimum value is ${min}.`;

        if (typeof max === "number" && value > max) return `Too high. Maximum value is ${max}.`;

        return null;
    }

    /** User date validation
     * @param value Date value to validate
     * @param isRequired Boolean if the field is required/mandatory
     * @param min Minimum date value
     * @param max Maximum date value
     * @returns Error message or null
     */
    private validateDate(value: string | Date, isRequired: boolean = false, min?: Date, max?: Date): string {
        const date = (typeof value === "string" || value instanceof Date) ? new Date(value as any) : null ,
            length = date ? date.toISOString().length : 0;

        if (isRequired && length === 0) return "Required.";

        if (min instanceof Date && date < min) return `Not valid. Minimum date is ${min.toLocaleString()}.`;
        
        if (max instanceof Date && date > max) return `Not valid. Maximum date is ${max.toLocaleString()}.`;

        return null;
    }

    /** User choice or multi-choice validation
     * @param values Array of values or singlar value
     * @param isRequired Boolean if the field is required/mandatory
     * @param min Minimum number of selections
     * @param max Maximum number of selections
     * @returns Error message or null
     */
    private validateChoice(values: any | any[], isRequired: boolean = false, min?: number, max?: number): string {
        if (typeof values === "string") return this.validateString(values, isRequired);
        if (typeof values === "number") return this.validateNumber(values, isRequired);

        const length = values instanceof Array ? values.length : 0;

        if (isRequired && length === 0) return "Required.";

        if (typeof min === "number" && length < min) return `Not enough. Select a miniumum of ${min} choice${min !== 1 ? "" : "s"}.`;
        
        if (typeof max === "number" && length > max) return `Too much. Select a maximum of ${max} choice${max !== 1 ? "" : "s"}.`;

        return null;
    }

    /** Creates the LookupData object (IAboutUsLookupItemValue)
     * @param values Lookup value or array of lookup values. Object(s) must contain 'ID' property
     * @param internalName Field InternalName used as the field's reference ID.
     * @returns LookupValue object that can beused for the Combo
     */
    private createValueState_LookupItem(values: any | any[], internalName: string): IAboutUsLookupItemValue {
        const field = this.getField_by_InternalName(internalName),
            lookupField = field.LookupField || null,
            value = {
                sp: field.AllowMultipleValues ? {results: []} : null,
                control: []
            };

        // return early if no value(s)
        if (!values) return value;

        // normalize values. make into array
        if (values instanceof Array === false) values = [values];

        // loop through each value
        for (let i = 0; i < values.length; i++) {
            const selected = values[i],
                // get the selected value key
                key = selected.Id || selected.ID || selected.id || selected.key || null,
                name = (lookupField ? selected[lookupField] : null) || selected.name || null;
                

            if (key) {
                if (field.AllowMultipleValues) {
                    value.sp.results.push(key);
                } else {
                    value.sp = key;
                }

                value.control.push({key: key, name: name});
            }
        }

        return value;
    }

    /** Creates the UserData object (IAboutUsUserValue).
     * Types of user data accepted:
     * - SP.FieldUser (single) = IUserInfoItem = {ID: number, Name: string}
     * - SP.FieldUser (multi) = IUserInfoItem[] = {results: [{ ID: number, Name: string }, ...]}
     * - PeoplePicker (SPFX PnP React Control) onChange = IPersonaProps[] = [{id: number, loginName: string, secondaryText: string}, ...]
     * - SP.User = ISiteUserProps = {Id: number, LoginName: string, Email: string, ...}
     * @param users SP UserInfoItem object or array of UserInfoItems
     * @param internalName Field InternalName used as the field's reference ID.
     * @returns UserData object that can be used for the PeoplePicker Control and SP REST API.
     * @example
     * // IAboutUsUserValue
     * valueState.value = {
     *     sp: { results: number[] } | number;  // user.Id(s)
     *     control: string[];   // [user.Login || user.Email, ...]
     * }
     */
    private createValueState_UserValue(users: any | any[], internalName: string): IAboutUsUserValue {
        const field = this.getField_by_InternalName(internalName),
            value = {
                sp: field.AllowMultipleValues ? {results: []} : null,
                control: []
            };

        // return early if users data is null
        if (!users) return value;

        // normalize users data. make all data an array of user infor (SPUSer, IUserInfoItem, IPersona)
        if (users instanceof Array === false) users = [users];

        //  loop through each user info
        for (let i = 0; i < users.length; i++) {
            const user = users[i],

                // get the value needed for the PeoplePicker control
                controlValue = (
                    // IPersonaProps
                    user.loginName ||
                    user.secondaryText ||
                    // ISiteUserProps
                    user.LoginName ||
                    user.Email ||
                    // IUserInfoItem
                    user.Name ||
                    // none of those properties exists
                    ""
                ).split("|").pop(),

                // get the value needed for SP REST API
                idValue = (
                    // IPersonaProps
                    user.id ||
                    // ISiteUserProps
                    user.Id ||
                    // IUserInfoItem
                    user.ID ||
                    // none of those properties exists
                    null
                );

            if (idValue) {
                if (field.AllowMultipleValues) {
                    value.sp.results.push(idValue);
                } else {
                    value.sp = idValue;
                }
            }
            if (controlValue) value.control.push(controlValue);
                
        }

        return value;
    }

    /** Initialize a new ValueState: {defaultValue, value, errorMessage}
     * @param id Field InternalName
     * @param value Initial or default value
     * @returns ValueState object
     */
    private initializeFieldValueState(value?, disabled: boolean = false): IAboutUsValueState {
        if (value === undefined) value = null;
        return {
            defaultValue: value,
            value: value,
            errorMessage: null,
            disabled: disabled
        };
    }

    /** Generates the IDropdownOption's array from array of string values.
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

    /** Parses SP's (";#") delimited string into an array of values
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

    /** Converts a string into a number
     * @param str String to convert
     * @returns Number or null
     */
    private stringToNumber(str: string | number): number {
        const num = parseFloat((str as string));
        return isNaN(num) ? null : num;
    }

    /** Get Field information
     * @param id Field InternalName of field to retrieve
     * @returns Field info
     */
    private getField_by_InternalName(InternalName: string): IDataFactoryFieldInfo {
        return find(this.props.list.fields, ["InternalName", InternalName]) || null;
    }

    /** Get the ValueState ({defaultValue, value, errorMessage}) for this field
     * @param InternalName Field InternalName for the ValueState object to retrieve
     * @returns ValueState object for Field ID passed
     */
    public getValueState_by_InternalName(InternalName: string): IAboutUsValueState {
        return (this.state.valueStates && Object.prototype.hasOwnProperty.call(this.state.valueStates, InternalName)) 
            ? this.state.valueStates[InternalName] 
            : null ;
    }

    /** Update the ValueState for a specific field. Updates the state and display. Pauses, to allow the state to update correctly.
     * @param InternalName Field InternalName for the ValueState object to set.
     * @param valueState New ValueState to set.
     */
    public async setValueState_for(InternalName: string, valueState: IAboutUsValueState): Promise<void> {
        this.setState({ "valueStates": {...this.state.valueStates, InternalName: valueState} });
        // setState needs just a momemt to update the state. 
        // not required but helps if you need to reference it right away.
        await this.sleep(0);
    }

    private getTextFromHTMLString(htmlString: string): string {
        this._div.innerHTML = htmlString || "";

        return this._div.innerText;
    }

    /** Pauses the script for a set amount of time.
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

    /** Prints our debug messages. Decorated console.info() or console.error() method.
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
//#endregion
}