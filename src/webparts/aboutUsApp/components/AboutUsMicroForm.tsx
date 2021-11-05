// About-Us micro forms (mini forms) for custom fields with complex data types
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { trim, assign } from 'lodash';
import * as strings from 'AboutUsAppWebPartStrings';

import { Dropdown,
    Label, 
    TextField,
    Checkbox,
    IDropdownOption,
    Link,
    Stack,
    IStackTokens,
    IStackStyles} from 'office-ui-fabric-react';
import CustomDialog from './CustomDialog';
import DataFactory, { IAboutUsMicroFormField } from './DataFactory';

//#region INTERFACES & ENUMS
export interface IAboutUsMicroFormValues {
    [key: string]: any;
}

export interface IMicroFormProps {
    fields: IAboutUsMicroFormField[];
    formValues: IAboutUsMicroFormValues;
    stateUpdated: (state: IMicroFormState)=>{};
    styles?: IStackStyles;
}

interface IMicroFormState extends IAboutUsMicroFormValues {
    errorMessage?: {[key: string]: string};
}
//#endregion

export class MicroForm extends React.Component<IMicroFormProps, IMicroFormState> {
//#region RENDER
    constructor(props) {
        super(props);

        const state = { errorMessage: {} };

        // need to create the state object with field keys and errorMessage keys.
        for (let i = 0; i < this.props.fields.length; i++) {
            const field = this.props.fields[i],
                key = field.internalName;

            // only create state for fields with form values
            if (key in this.props.formValues) {
                const value = this.props.formValues[key];

                state[key] = value;
                state.errorMessage[key] = this.fieldRequiredMessage(field, value);
            }
        }

        // save init state
        this.state = state;
    }

    public render(): React.ReactElement<IMicroFormProps> {
        const tokens: IStackTokens = {
                childrenGap: 10
            };

        return (
            <Stack tokens={tokens} styles={this.props.styles}>
                { this.props.fields.map(field => this.MicroFormControl(field) ) }
            </Stack>
        );
    }

//#endregion

//#region CONTROLS
    private MicroFormControl(field: IAboutUsMicroFormField): React.ReactElement {
        const fieldId = field.id || `microForm-${field.internalName}`;

        // don't render any hidden fields
        if (field.hidden === true) return;

        // return control based on type
        switch (field.type) {
            case "checkbox":
                return <Checkbox
                    id={fieldId}
                    label={field.label || null}
                    ariaLabel={field.label + " checkbox"}
                    placeholder={field.placeholder || ""}
                    defaultValue={this.state[field.internalName]}
                    onChange={ (evt, newValue) => { this.onChange_control(field, newValue); } }
                    disabled={field.disabled || false}
                    styles={field.styles || {}}
                    className={field.className || ""} />;
        
            case "dropdown":
                return <Dropdown
                    id={fieldId}
                    label={field.label || null}
                    ariaLabel={field.label + " dropdown"}
                    placeholder={field.placeholder || null}
                    defaultSelectedKey={this.state[field.internalName]}
                    onChange={ (evt, newValue) => { this.onChange_control(field, newValue.key); } }
                    options={field.options}
                    required={field.required || false}
                    disabled={field.disabled || false}
                    styles={field.styles || {}}
                    className={field.className || ""} />;
        
            case "multiselect":
                return <Dropdown
                    id={fieldId}
                    label={field.label || null}
                    ariaLabel={field.label + " multi-select dropdown"}
                    placeholder={field.placeholder || null}
                    multiSelect={true}
                    defaultSelectedKeys={this.state[field.internalName]}
                    onChange={ (evt, newValue) => { this.onChange_miltiselect(field, newValue); } }
                    options={field.options}
                    required={field.required || false}
                    disabled={field.disabled || false}
                    styles={field.styles || {}}
                    className={field.className || ""} />;

            case "link":
                return <Link
                    id={fieldId}
                    title={field.placeholder || field.label || ""}
                    href={this.state[field.internalName]}
                    target={field.target || "_blank"}
                    disabled={field.disabled || false}
                    styles={field.styles || {}}
                    className={field.className || ""}
                    >{field.label || ""}</Link>;

            case "label":
                return <Label
                    htmlFor={fieldId || this.state[field.internalName] || null}
                    required={field.required || false}
                    disabled={field.disabled || false}
                    styles={field.styles || {}}
                    className={field.className || ""}
                    >{field.label || ""}</Label>;
                            
            default:    // assumes it's an textbox type
                return <TextField
                    id={fieldId}
                    type={field.type}
                    label={field.label || null}
                    ariaLabel={field.label + " textbox"}
                    placeholder={field.placeholder || field.label || ""}
                    description={field.description || null}
                    errorMessage={ this.state.errorMessage[field.internalName] }
                    multiline={field.multiline || false}
                    defaultValue={this.state[field.internalName]}
                    onChange={ (evt, newValue) => { this.onChange_control(field, newValue); } }
                    required={field.required || false}
                    disabled={field.disabled || false}
                    styles={field.styles || { root: {"min-width": "400px"} }}
                    className={field.className || ""} />;

        }
    }

    private onChange_control(field: IAboutUsMicroFormField, newValue: any) {
        if (field.internalName in this.state) {
            this.setState({[field.internalName]: newValue});

            // required?
            const errorMessage = this.fieldRequiredMessage(field, newValue);
            this.setState({errorMessage: {...this.state.errorMessage, [field.internalName]: errorMessage}});

            setTimeout(()=>{this.props.stateUpdated(this.state);},0);
       }
    }

    private onChange_miltiselect(field: IAboutUsMicroFormField, newValue: IDropdownOption) {
        if (field.internalName in this.state) {
            let value: any[] = this.state[field.internalName];

            // make sure form value is an array
            if (!(value instanceof Array)) value = [];

            const key = newValue.key,
                ndx = value.indexOf(key);

            // add or remove from form values
            if (newValue.selected) {
                if (ndx === -1) value.push(key);

            } else {
                if (ndx > -1) value.splice(ndx, 1);
            }

            // update value
            this.onChange_control(field, value);
        }
    }


//#endregion

//#region HELPERS
    private fieldRequiredMessage(field: IAboutUsMicroFormField, value): string {
        let errorMessage = "";

        if (field.required === true && MicroForm.isValueEmpty(value)) {
            errorMessage = "This field is required.";
        }

        return errorMessage;
    }

    public static isValueEmpty(value): boolean {
        // if null
        if (value === null || value === undefined) return true;

        // array
        if (value instanceof Array && value.length === 0) return true;

        // string
        if (typeof value === "string" && trim(value).length === 0) return true;

        // number
        if (typeof value === "number" && (value + "").length === 0) return true;

        // object
        if (typeof value === "object" && Object.keys(value).length === 0) return true;

        return false;
    }
//#endregion
}

export default class AboutUsMicroForm extends CustomDialog {
//#region PROPERTIES
    private formValues_: IAboutUsMicroFormValues;
    private formFields_: IAboutUsMicroFormField[];
    private formStyles_: IStackStyles;
//#endregion

//#region CONSTRUCTOR
    constructor(title: string, fields: IAboutUsMicroFormField[], defaultValue?: IAboutUsMicroFormValues, formStyles?: IStackStyles) {
        super(title, true);

        this.formFields_ = fields;
        this.formValues_ = AboutUsMicroForm.initFormValues(fields, defaultValue);
        this.formStyles_ = formStyles;

        this.generateForm(fields, this.formValues_);
    }

    public async show(): Promise<any> {
        const dialog = this;
        let okClicked = false;

        dialog.AddCancelAction("Cancel", {tabIndex: 0}).AddSubmit(evt => {
            okClicked = true;
            dialog.close();
            return false;
        });

        await super.show();

        return (okClicked) ? this.formValues_ : null ;
    }
//#endregion

//#region FORM FIELDS
    private generateForm(fields: IAboutUsMicroFormField[], formValues: IAboutUsMicroFormValues) {
        const root = this.body;

        ReactDOM.render(<MicroForm
            fields={fields}
            formValues={formValues}
            stateUpdated={this.microForm_stateUpdated.bind(this)}
            styles={this.formStyles_} />, root);
    }

    private microForm_stateUpdated(state: IMicroFormState) {
        let formValues = assign({}, state);
        delete formValues.errorMessage;
        this.formValues_ = formValues;
    }
//#endregion

//#region HELPERS
    /** Initializes the form values. If default values was passed or uses the 
     * @param fields Form fields
     * @param formValues Default value object that should match the form field's internal name
     * @returns Form value object with form field's internal name.
     */
    public static initFormValues(fields: IAboutUsMicroFormField[], formValues: IAboutUsMicroFormValues = null): any {
        const skipType = ["label", "link"],
            values = {};

        fields.forEach(field => {
            const key = field.internalName;

            // only fields that are in the micro form template are used as the value object.
            if (skipType.indexOf(field.type) === -1) {
                // if form values were passed, use as the starting value (assumes this is an edit form),
                //  else use the field default (assumes this is a new form).
                const value = (!formValues) ? field.defaultValue : formValues[key] ;
                values[key] = (value === undefined) ? null : value ;
            }
        });

        return values;
    }
//#endregion
}