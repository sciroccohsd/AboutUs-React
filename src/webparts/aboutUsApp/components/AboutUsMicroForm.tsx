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
    IStackStyles,
    ILinkProps,
    ITextField,
    ITextFieldProps} from 'office-ui-fabric-react';
import CustomDialog from './CustomDialog';
import DataFactory, { IAboutUsMicroFormField } from './DataFactory';
import { FilePicker, IFilePickerProps, IFilePickerResult } from '@pnp/spfx-controls-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { arrayListFormat, DEBUG_NOTRACE, IAboutUsAppWebPartProps } from '../AboutUsAppWebPart';
import { IItem } from '@pnp/sp/items';

//#region INTERFACES & ENUMS
export interface IAboutUsMicroFormValues {
    [key: string]: any;
}

export interface IMicroFormProps {
    ctx: WebPartContext;   // required for file picker
    properties: IAboutUsAppWebPartProps;
    fields: IAboutUsMicroFormField[];
    formValues: IAboutUsMicroFormValues;
    stateUpdated: (state: IMicroFormState)=>{};
    styles?: IStackStyles;
    resourceFolder?: string;    // for file picker. upload folder name inside external repo library.
}

interface IMicroFormState extends IAboutUsMicroFormValues {
    errorMessage?: Record<string, string>;
    defaultFolderAbsolutePath?: string[];
}
//#endregion

export class MicroForm extends React.Component<IMicroFormProps, IMicroFormState> {
//#region RENDER
    constructor(props) {
        super(props);

        const state = {
            errorMessage: {},
            defaultFolderAbsolutePath: [this.props.ctx.pageContext.web.absoluteUrl]
        };

        if (this.props.properties.externalRepo) {
            state.defaultFolderAbsolutePath.push(this.props.properties.externalRepo.split("/").pop());
        }

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

    public async componentDidMount() {
        // ensure folder if resource folder was passed
        if (this.props.properties.externalRepo && this.props.resourceFolder) {
            const exists = await DataFactory.ensureFolder(this.props.properties.externalRepo, this.props.resourceFolder);
            if (exists) {
                const defaultFolderAbsolutePath = [...this.state.defaultFolderAbsolutePath];
                defaultFolderAbsolutePath.push(this.props.resourceFolder);
                this.setState({"defaultFolderAbsolutePath": defaultFolderAbsolutePath});
            }
        }
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
                    className={field.className || ""} />;

            case "label":
                return <Label
                    htmlFor={fieldId || this.state[field.internalName] || null}
                    required={field.required || false}
                    disabled={field.disabled || false}
                    styles={field.styles || {}}
                    className={field.className || ""}
                    >{field.label || ""}</Label>;
                            
            default:    // assumes it's an textbox type
                return this.TextFieldControl(fieldId, field);

        }
    }

    private TextFieldControl(fieldId: string, field: IAboutUsMicroFormField): React.ReactElement {
        const props: ITextFieldProps = {
                id: fieldId,
                type: field.type,
                label: field.label || null,
                ariaLabel: field.label + " textbox",
                placeholder: field.placeholder || field.label || "",
                description: field.description || null,
                errorMessage:  this.state.errorMessage[field.internalName] ,
                multiline: field.multiline || false,
                value: this.state[field.internalName],
                onChange:  (evt, newValue) => { this.onChange_control(field, newValue); },
                required: field.required || false,
                disabled: field.disabled || false,
                styles: field.styles || { root: {"min-width": "400px"} },
                className: field.className || ""
            },
            picker_onChange = (results: (IFilePickerResult | IFilePickerResult[])) => {
                DEBUG_NOTRACE("MicroFormControl > TextFieldControl > FilePicker:", results);

                const result = (results && results instanceof Array && results.length > 0) ? results[0] : results as IFilePickerResult;

                const url = result.fileAbsoluteUrl;

                this.onChange_control(field, url);
            },
            filePickerProps: IFilePickerProps = (field.filePickerProps)
                ? assign({
                    context: this.props.ctx,
                    onSave: (items: IFilePickerResult[]) => { this.filePicker_onSave(items, field); },
                    onChange: (items: IFilePickerResult[]) => { this.filePicker_onChange(items, field); },
                    buttonLabel: "Select a file",
                    buttonIcon: "FileImage",
                    defaultFolderAbsolutePath: this.state.defaultFolderAbsolutePath.join("/")
                }, field.filePickerProps) as IFilePickerProps
                : null ;
        
        return (
            <div>
                { React.createElement(TextField, props) }
                { filePickerProps ? React.createElement(FilePicker, filePickerProps) : null }
            </div>
            
        );
    }

    //#region FILE PICKER
    private updateFilePickerField(item: IFilePickerResult, field: IAboutUsMicroFormField) {
        const url = item.fileAbsoluteUrl;

        // if no url, item  may need to be uploaded.
        if (!url) return;

        this.onChange_control(field, url);
    }

    private async uploadItem(item: IFilePickerResult, field: IAboutUsMicroFormField): Promise<IItem> {
        const file = await item.downloadFileContent();
        let uploadedItem: IItem = null;

        // invalid file type
        if (field.filePickerProps.accepts && field.filePickerProps.accepts.length > 0) {
            if (field.filePickerProps.accepts.indexOf(file.name.toLowerCase().split(".").pop()) === -1) {
                // not accepted file type
                await CustomDialog.alert(
                    `Unable to upload file. Allowed types: ${arrayListFormat(field.filePickerProps.accepts, ", ", " or ")}`,
                    "Invalid file type");

                return uploadedItem;
            }
        }

        // only upload files that are 5MBs or lower
        if (file.size <= 5500000) {
            uploadedItem = await DataFactory.uploadFile(this.props.properties.externalRepo, file, this.props.resourceFolder);
                
        } else {
            // too large
            await CustomDialog.alert("Unable to upload file. Image size must be less than 3 MBs.", "File too large");
        }

        return uploadedItem;
    }

    private filePicker_onChange(items: IFilePickerResult[], field: IAboutUsMicroFormField) {
        // the URL field can only accept one item
        const item = (items && items.length > 0) ? items[0] : null;
        if (item) this.updateFilePickerField(item, field);
    }

    private async filePicker_onSave(items: IFilePickerResult[], field: IAboutUsMicroFormField) {
        // external repo not set up
        if (!this.props.properties.externalRepo) {
            return await CustomDialog.alert("Administrators haven not configured the external repository.", "Unable to upload file");
        }

        // no file was selected
        if (!items || items.length === 0) return;

        // parse selected file
        items.forEach( async (item) => {
            // if item already has a URL, item is already uploaded.
            if (item.fileAbsoluteUrl) return this.updateFilePickerField(item, field);

            // upload this file
            const uploadItem = await this.uploadItem(item, field);
            if (uploadItem) {
                item.fileAbsoluteUrl = new URL((uploadItem as Record<string, any>).ServerRelativeUrl, location.origin).href;
                item.fileName = (uploadItem as Record<string, any>).Name;

                this.updateFilePickerField(item, field);
                
            } else {
                // something went wrong with the upload
                await CustomDialog.alert("Unable to upload file. View console for more details.", "ERROR!");
            }

        });
    }
    //#endregion

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
    private ctx: WebPartContext;
    private properties: IAboutUsAppWebPartProps;
    private formValues: IAboutUsMicroFormValues;
    private formStyles: IStackStyles;
    private resourceFolder: string;
//#endregion

//#region CONSTRUCTOR
    constructor(
            ctx: WebPartContext,
            properties: IAboutUsAppWebPartProps,
            title: string,
            fields: IAboutUsMicroFormField[],
            defaultValue?: IAboutUsMicroFormValues,
            formStyles?: IStackStyles,
            resourceFolder?: string
        ) {

        super(title, true);

        this.ctx = ctx;
        this.properties = properties;
        this.formValues = AboutUsMicroForm.initFormValues(fields, defaultValue);
        this.formStyles = formStyles;
        this.resourceFolder = resourceFolder;

        this.generateForm(fields, this.formValues);
    }

    public async show(): Promise<any> {
        const dialog = this;
        let okClicked = false;

        dialog.AddCancelAction("Cancel", {tabIndex: 0}).AddSubmit(evt => {
            okClicked = true;
            dialog.close();
            return false;
        }, "OK");

        await super.show();

        return (okClicked) ? this.formValues : null ;
    }
//#endregion

//#region FORM FIELDS
    private generateForm(fields: IAboutUsMicroFormField[], formValues: IAboutUsMicroFormValues) {
        const root = this.body;

        ReactDOM.render(<MicroForm
            ctx={this.ctx}
            properties={this.properties}
            fields={fields}
            formValues={formValues}
            stateUpdated={this.microForm_stateUpdated.bind(this)}
            styles={this.formStyles}
            resourceFolder={this.resourceFolder} />, root);
    }

    private microForm_stateUpdated(state: IMicroFormState) {
        let formValues = assign({}, state);
        delete formValues.errorMessage;
        this.formValues = formValues;
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