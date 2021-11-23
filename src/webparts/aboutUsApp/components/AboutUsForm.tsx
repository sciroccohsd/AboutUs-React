// About-Us custom new and edit form
import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { find, trim, escape, assign } from 'lodash';
import * as strings from 'AboutUsAppWebPartStrings';
import { WebPartContext, BaseWebPartContext } from '@microsoft/sp-webpart-base';

import DataFactory, { IDataFactoryFieldInfo, IFieldUrlValue, IUserInfoItem } from './DataFactory';
import CustomDialog from './CustomDialog';
import * as FormControls from './FormControls';
import * as AboutUsDisplay from './AboutUsDisplay';
import AboutUsMicroForm from './AboutUsMicroForm';

import { CommandBarButton, 
    ICommandBarItemProps, 
    IContextualMenuItem, 
    IDropdownOption,
    IStackStyles,
    IStackTokens,
    ITextFieldProps, 
    MessageBar, 
    MessageBarType, 
    Stack,
    } from 'office-ui-fabric-react';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react';
import { IItemAddResult, IItemUpdateResult, _Items } from '@pnp/sp/items/types';
import { PermissionKind } from '@pnp/sp/security';
import "@pnp/sp/security";
import { IAboutUsAppWebPartProps, IAboutUsAppFieldOption, isInRange_numDays, sleep, LOG } from '../AboutUsAppWebPart';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';

//#region INTERFACES, TYPES & ENUMS
export interface IAboutUsFormProps {
    ctx: WebPartContext;
    list: DataFactory;
    form: "new" | "edit";
    properties: IAboutUsAppWebPartProps;
    history?: History;
    itemId?: number;
}

// Form field's current value, default/initial value, error status, disabled status
export interface IAboutUsValueState {
    defaultValue: any;  // starting or default value
    value: any;         // current value on form
    errorMessage: string;  // error message if any, else null or "";
    disabled: boolean;  // field is disabled
}

export interface IAboutUsMultiChoiceItemValue {
    sp: {results: any[]} | any;
    control: any[] | any;
}

export interface IAboutUsUserValue {
    sp: { results: number[] } | number;  // user.Id(s)
    control: string[];   // [user.Login || user.Email, ...]
}

export type TAboutUsComplexData = Record<string, any>;
export interface IAboutUsComplexValue {
    sp: string;  // JSON.stringify(data)
    control: TAboutUsComplexData | TAboutUsComplexData[];  // JSON.parse(data)
}

export interface IAboutUsKeywordsValue {
    sp: string;
    control: string[];
}

enum DISPLAY_STATE {
    "loading",
    "invalid",
    "ready",
    "saving"
}

interface IAboutUsFormState {
    display: DISPLAY_STATE;     // form display's state

    isAdmin: boolean;  // can add/edit/delete controlled fields and list items

    canSaveForm: boolean;
    canCancelForm: boolean;
    canDeleteItem: boolean;

    isProcessingForm: boolean;

    errorMessage: string;
}
//#endregion


export default class AboutUsForms extends React.Component<IAboutUsFormProps, IAboutUsFormState & any, {}> {
//#region PROPERTIES
    private baseComponentContext: BaseWebPartContext = null;
    private listItem: Record<string, any> = null;
    private fieldsThatHaveBeenModified: string[] = [];   // list of internal field names that have been modified/updated
    private user: ISiteUserInfo = null;

    private _htmlNode = document.createElement("div");

    private valueStateKeyPrefix = "valueState_";
//#endregion

//#region RENDER
    constructor(props) {
        super(props);

        // state default values
        this.state = {
            display: DISPLAY_STATE.loading, // display loading message?

            isAdmin: false,         // user can add & delete items

            canSaveForm: false,     // user can add or edit items
            canCancelForm: true,
            canDeleteItem: false,   // user can delete itemss

            isProcessingForm: false,    // disables form buttons if processing a request

            errorMessage: null      // global error message.
        };

    }
    
    public render(): React.ReactElement<IAboutUsFormProps> {
        this.baseComponentContext = this.props.ctx as any;
        return (
            <div className={styles.form}>
                { this.state.display === DISPLAY_STATE.ready ? 
                    <form>
                        <h3 className={styles.formHeader}>
                            { (this.props.form === "new") 
                                ? "New About-Us Entry:" 
                                : `Editing ${ this.listItem.Title }:`
                            }
                        </h3>
                        { (this.props.form) ? this.formFields() : null }
                    </form>
                : 
                    <div>
                        { this.state.display === DISPLAY_STATE.loading ? <FormControls.LoadingSpinner/> : null }
                        { this.state.display === DISPLAY_STATE.invalid ? <this.InvalidItem/> : null }
                    </div>
                }
            </div>
        );
    }

    public async componentDidMount() {
        // let _item = null;
        await this.setCurrentUserFlags();

        if (this.props.form === "new") {
            this.setState({"display": DISPLAY_STATE.ready});
            return;

        } else if (this.props.form === "edit") {
            // edit form: need to get list item data
            this.setState({"display": DISPLAY_STATE.loading});

            this.listItem = await this.props.list.getItemById_expandFields(this.props.itemId);

            // debug
            LOG("this.listItem:", this.listItem);

            // check to see if an item exists
            if (this.listItem !== null) {
                this.setState({"display": DISPLAY_STATE.ready});

            } else {
                // item does not exist or ID/Office is invalid
                this.setState({"display": DISPLAY_STATE.invalid});
            }
        }
    }

    private InvalidItem(): React.ReactElement {
        return (
            <div className={styles.aboutUsApp}>
                <h3>Invalid item ID or Office Symbol</h3>
                <p>Unable to retrieve About-Us item. 
                    Please check to ensure the item ID is correct. 
                    Please contact the administrators [ADD_ADMIN_MAILTO] if you have any question.
                </p>
            </div>
        );
    }

    private ErrorMessage(): React.ReactElement {
        return (
            <div>
                <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
            </div>
        );
    }
//#endregion

//#region FORM FIELDS
    private formFields(): React.ReactElement {
        let elements: React.ReactElement[] = [];

        this.props.list.fields.forEach( field => {
            try{
                const defaultValue = (this.props.form === "edit") 
                    ? this.getListItemData_for(field.InternalName) 
                    : field.DefaultValue,
                    valueState = this.ensureValueState_for(field.InternalName, defaultValue);

                let element = this.createFieldControl(field, valueState);
                if (element) elements.push(element);

            } catch(er) {
                LOG("ERROR: form()", field, er);
            }
        });

        return (
            <div>
                { this.NewEditCommandBar() }
                { this.state.errorMessage ? this.ErrorMessage() : null }
                <Stack tokens={{ childrenGap: 10}} className={styles.formFieldsContainer}>{ elements }</Stack>
                { this.state.errorMessage ? this.ErrorMessage() : null }
                { this.NewEditCommandBar() }
            </div>
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

        // custom 'ComplexData' form controls
        const displayControls: Record<string, React.ComponentClass<AboutUsDisplay.IAboutUsComplexDataDisplayProps>> = {
            "Tasks": AboutUsDisplay.TasksDisplay,
            "Bios": AboutUsDisplay.BiosDisplay,
            "Links": AboutUsDisplay.LinksDisplay,
            "SOP": AboutUsDisplay.SOPDisplay,
            "Contacts": AboutUsDisplay.ContactsDisplay
        };

        switch (field.InternalName) {
            case "Tasks":
            case "Bios":
            case "Links":
            case "SOP":
            case "Contacts":
                // use the default ComplexData type controls and display
                // ComplexData type: Uses a micro-form (from AboutUsListTemplate.json) to gather data.
                //  Pass in the AboutUsDisplay control to render the items differently
                if (!valueState.value || typeof valueState.value === "string") {
                    valueState.value = this.createValueState_ComplexData(valueState.value);
                }

                // get the display control for this specific
                const displayControl = displayControls[field.InternalName];

                // create custom complex data control
                if (displayControl) {
                    return this.customFieldComplexData(field, valueState, displayControl);
                } else {
                    LOG(`createFieldControl() > Custom Form Control: Could not find the AboutUsDisplay \
                        control type for ${field.InternalName}. Used default rendering instead.`, field);
                }

                break;

            case "Keywords":
                if (!valueState.value || typeof valueState.value === "string") {
                    valueState.value = this.createValueState_Keywords(valueState.value);
                }

                return this.customFieldKeywords(field, valueState);

            case "Logo":
                // url field with file picker
                return this.customFieldLogo(field, valueState);

            case "Validated": 
                return (this.props.form === "edit") ? this.customFieldValidated(field, valueState) : null ;

            case "BroadcastDate":
                return null;

            case "ValidatedBy":
                valueState.value = this.createValueState_UserValue(valueState.value, field.InternalName);
                return null;
        }
        
        // default form controls
        switch (field["odata.type"]) {
            case "SP.FieldText":
                return this.spFieldText(field, valueState);

            case "SP.FieldNumber":
                valueState.value = this.stringToNumber(valueState.value);
                return this.spFieldNumber(field, valueState);

            case "SP.FieldMultiLineText":
                if (field.RichText === true) {
                    // is multiline rich text  field
                    return this.spRichText(field, valueState);
                } else {
                    //  basic multiline field
                    return this.spFieldText(field, valueState);
                }
        
            case "SP.FieldChoice":
                return this.spFieldChoice(field, valueState);
                    
            case "SP.FieldMultiChoice":
                valueState.value = this.createValueState_MultiChoice( valueState.value );
                return this.spFieldMultiChoice(field, valueState);

            case "SP.FieldLookup":
                // valueState.value = selectedId || { results: [selectedId, ...] }
                valueState.value = this.createValueState_LookupItem(valueState.value,  field.InternalName);
                return this.spFieldLookup(field, valueState);

            case "SP.FieldUrl":
                /* 
                valueState.value = { 
                    _metadata: {type: "SP.FieldUrlValue"},
                    Url: string,
                    Description: string 
                }
                */
                return this.spFieldUrl(field, valueState);

            case "SP.FieldDateTime":
                // valueState.value = string ISO8601 date/time
                return this.spFieldDateTime(field, valueState);

            case "SP.FieldUser":
                /*
                valueState.value = {
                    sp: { results: number[] } | number;  // user.Id(s)
                    control: string[];   // [user.Login || user.Email, ...]
                }
                */
                valueState.value = this.createValueState_UserValue(valueState.value, field.InternalName);
                return this.spFieldUser(field, valueState);

            default:
                LOG("Unhandled field control:", field["odata.type"], field.InternalName, field);
                return null;
        }
    }

    // SPECIFIC CONTROLS
    /** Create a textbox control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private spFieldText(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            isMultiline = field["odata.type"] === "SP.FieldMultiLineText",
            props: ITextFieldProps = {
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IRichTextControlProps = {
                // RichText control doesn't display value if "isEditMode" is true.  isEditMode = !isDisabled;
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: ITextFieldProps = {
                type: "number",
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IDateTimeControlProps = {
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
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
                //onGetErrorMessage: (newDate) => {LOG("spFieldDateTime > onGetErrorMessage", newDate); return "";},

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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IDropdownControlProps = {
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IDropdownControlProps = {
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
                label: field.Title,
                placeholder: "Select an option",
                description: field.Description,

                multiSelect: true,
                options: this.generateDropDownOptions(field.Choices),
                defaultSelectedKeys: valueState.value.control,

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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IListItemPickerControlProps = {
            // props: FormControls.IComboBoxListItemPickerControlProps = {
                disabled: valueState.disabled,
                required: field.Required || fieldOption.required,
                label: field.Title,
                description: field.Description,

                // ListItemPickerControl properties
                placeholder: "Type to search...",
                context: this.baseComponentContext,
                enableDefaultSuggestions: true,

                // ComboBoxListItemPickerControl properties
                // spHttpClient: this.props.ctx.spHttpClient as any,
                // multiSelect:  field.AllowMultipleValues,

                // shared properties
                itemLimit: field.AllowMultipleValues ? 99 : 1,
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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IUrlControlProps = {
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
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
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props: FormControls.IPeoplePickerControlProps = {
                context: this.baseComponentContext,
                webAbsoluteUrl: this.props.ctx.pageContext.web.absoluteUrl,
                disabled: valueState.disabled,
                required: field.Required || fieldOption.required,
                label: field.Title,
                placeholder: `Select ${ field.Title }...`,
                description: field.Description,

                defaultSelectedUsers: valueState.value.control,
                ensureUser: true,   // if false, user.id will be a string & not the user ID number
                personSelectionLimit: field.AllowMultipleValues ? 99 : 1,

                errorMessage: valueState.errorMessage,

                onChange: users => this.peoplePicker_onChange((users as IPersonaProps[]), field.InternalName)
            };

        return React.createElement(FormControls.PeoplePickerControl, props);
    }

    /** Create a Custom Complex Data control. Array of JSON data that normally requires a micro-form.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */
    private customFieldComplexData(
        field: IDataFactoryFieldInfo,
        valueState: IAboutUsValueState,
        displayControl: React.ComponentClass<any>): React.ReactElement {

        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props = {
                disabled: valueState.disabled,
                required: field.Required || fieldOption.required,
                label: field.Title,
                description: field.Description,
                values: valueState.value.control,
                errorMessage: valueState.errorMessage,

                displayControl: displayControl,

                properties: this.props.properties,
                showEditControls: true,
                onAdd: () => { this.complexData_onAdd(field.InternalName); },
                onEdit: (ndx: number) => { this.complexData_onEdit(field.InternalName, ndx); },
                onOrderChange: (oldIndex: number, newIndex: number) => { this.reorderValueState_ArrayData(field.InternalName, oldIndex, newIndex); },
                onDelete: (ndx: number)=>{ this.arrayData_onDelete(field.InternalName, ndx); },
                extraButtons: (field.InternalName === "Bios") ? this.commandbarBroadcastButton.bind(this) : null
            };

        return React.createElement(FormControls.CustomControlComplexData, props);
    }

    /** Create a Custom 'Keyword' control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */    
    private customFieldKeywords(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            props = {
                disabled: valueState.disabled,
                required: field.Required || fieldOption.required,
                label: field.Title,
                description: field.Description,
                values: valueState.value.control,
                errorMessage: valueState.errorMessage,

                showEditControls: true,
                onAdd: (value) => { return this.arrayData_onAdd(field.InternalName, value); },
                onOrderChange: (oldIndex: number, newIndex: number) => {
                    this.reorderValueState_ArrayData(field.InternalName, oldIndex, newIndex);
                },
                onDelete: (ndx: number)=>{ this.arrayData_onDelete(field.InternalName, ndx); }
            };

        return React.createElement(FormControls.CustomControlKeywords, props);
    }

    /** Create a Custom 'Validated' control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */    
     private customFieldValidated(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const validatedBy = "ValidatedBy",
            validatedByValue = (this.props.form === "edit") ? this.getListItemData_for(validatedBy) : null,
            validatedByState = this.ensureValueState_for(validatedBy, {
                control: validatedByValue,
                sp: (validatedByValue) ? validatedByValue.ID : null
            }),
            validated = valueState.value,
            props: AboutUsDisplay.IPageValidationDisplayProps = {
                properties: this.props.properties,
                validated: (validated) ? new Date(validated) : null,
                validatedBy: validatedByState.value.control,
                showButton: !valueState.disabled,
                onValidate: async () => { await this.validated_onClick(field.InternalName, validatedBy); }
            };

        return <div className={styles.fieldWrapper}>
            <div className={styles.validatedContainer}>
                { React.createElement(AboutUsDisplay.PageValidationDisplay, props) }
            </div>
        </div>;
    }

    /** Create a Custom 'Logo' control.
     * @param field Field information
     * @param valueState ValueState object for this field
     * @returns Form element with with label, required asterisk, field, description, & error text elements.
     */ 
    private customFieldLogo(field: IDataFactoryFieldInfo, valueState: IAboutUsValueState): React.ReactElement {
        const fieldOption = this.getFieldWebPartOptions_by_InternalName(field.InternalName),
            titleValue = this.getValueState_by_InternalName("Title"),
            siteUrl = this.props.ctx.pageContext.web.absoluteUrl,
            externalRepo = this.props.properties.externalRepo || "",
            props: FormControls.IUrlControlProps = {
                disabled: valueState.disabled,
                id: field.InternalName,
                required: field.Required || fieldOption.required,
                label: field.Title,
                description: field.Description,

                defaultValue: valueState.value,

                errorMessage: valueState.errorMessage,

                onChange: this.urlfield_onChange.bind(this),

                //uploadFolder: (externalRepo) ? [externalRepo, folderName].join("/") : "",
                externalRepo: externalRepo,
                folderName: trim(titleValue.value),
                filePickerProps: {
                    context: this.props.ctx,

                    // optional
                    label: "Or select the logo from this site:",
                    buttonLabel: "Select a logo",
                    accepts: ["jpg", "jpeg", "png", "svg", "ico"],
                    hideWebSearchTab: true,
                    hideStockImages: true,
                    hideOrganisationalAssetTab: true,
                    hideOneDriveTab: true,
                    hideLocalUploadTab: true,
                    includePageLibraries: false
                }
            };

        return React.createElement(FormControls.UrlControl, props);
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
                isNumber = field.TypeAsString === "Number",
                value = isNumber ? this.stringToNumber(evt.target.value) : evt.target.value,
                valueState = this.getValueState_by_InternalName(internalName);

            if (valueState === null) return;

            valueState.value = isNumber ? value : trim(value as string) || null;
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, value);

            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! input_change():", evt.target.id, er);
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
            const valueState = this.getValueState_by_InternalName(internalName),
                contentText = this.getTextFromHTMLString(value),
                contentLength = trim(contentText || "").length;

            if (valueState === null) return;

            valueState.value = contentLength > 0 ? value : null;
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, value);

            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! richtext_change():", internalName, er);
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
                // is multi-choice: value returned is what changed, not what's selected
                // need to update the existing value with what changed
                if (value.selected) {
                    if (valueState.value.sp === null) valueState.value.sp = {results: []};
                    if (valueState.value.control === null) valueState.value.control = [];

                    // new value added. need to check to see if this was called multiple times
                    if (valueState.value.sp.results.indexOf(value.key) === -1) {
                        valueState.value.sp.results.push(value.key);
                        valueState.value.control.push(value.key);
                    }

                } else {
                    if (valueState.value.sp !== null && valueState.value.control !== null) {
                        // remove selected value
                        const ndx = valueState.value.control.indexOf(value.key);
                        if (ndx > -1) {
                            valueState.value.sp.results.splice(ndx, 1);
                            valueState.value.control.splice(ndx, 1);
                        }
                    }

                }

            } else {
                // is single-choice: value returned is the new value
                valueState.value = value.key;
            }
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, valueState.value);

            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! dropdown_change():", evt.target.id, er);
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
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, value);

            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! lookup_onSelectedItem():", internalName, er);
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
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, value);

            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! urlfield_onChange():", internalName, er);
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

            valueState.value = newDate instanceof Date ? newDate.toISOString() : null;
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, newDate);

            // must setState for date fields, if not, the date value disappears from the UI
            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! datetime_onChange():", internalName, er);
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
            this.fieldWasModified(internalName);

            // validate / set errorMessage
            valueState.errorMessage = this.validateFieldValue(internalName, value);

            this.setValueState_for(internalName, valueState);

        } catch(er) {
            LOG("ERROR! peoplePicker_onChange():", internalName, er);
        }
    }

    /** Complex data Add Button event handler.
     * Opens a micro form to populate the array of values
     */
    private complexData_onAdd(internalName: string) {
        // field uses a mico-form to help gather data.
        // Requires a "fieldMicroForms" entry in the "AboutUsListTemplate.json"
        const microForm = this.generateMicroForm(internalName);

        // make sure there is a micro form for this field
        if (!microForm) return;

        // show form and wait for response
        microForm.show().then(value => {
            this.updateValueState_ArrayData(internalName, value);
        });
    }

    /** Complex data Edit item event handler
     * @param ndx Index of item that was clicked
     */
    private complexData_onEdit(internalName: string, ndx: number) {
        // field uses a mico-form to help gather data.
        // Requires a "fieldMicroForms" entry in the "AboutUsListTemplate.json"
        const valueState = this.getValueState_by_InternalName(internalName),
            microForm = this.generateMicroForm(internalName, valueState.value.control[ndx]);

        // make sure there is a micro form for this field
        if (!microForm) return;

        // show form and wait for response
        microForm.show().then(value => {
            this.updateValueState_ArrayData(internalName, value, ndx);
        });
    }

    /** Add value to array data
     * @param internalName Field InternalName used as the fields reference ID
     * @param value New value to to the array
     * @returns Input:Textbox value state
     */
    private arrayData_onAdd(internalName: string, value: string): FormControls.ICustomControlKeywordsState {
        const valueState = this.getValueState_by_InternalName(internalName),
            state = {
                value: ""
            };

        // if null or empty, don't save
        if (!value || trim(value).length === 0) {
            state.value = value;
            valueState.errorMessage = "Empty value!";
            this.setValueState_for(internalName, valueState);
            return state;
        }

        value = trim(value);
        const lcValue = value.toLowerCase();

        // if duplicate, don't save
        for (let i = 0; i < valueState.value.control.length; i++) {
            const existingValue = trim(valueState.value.control[i]).toLowerCase();
            if (existingValue === lcValue) {
                state.value = value;
                valueState.errorMessage = "Duplicate value!";
                this.setValueState_for(internalName, valueState);
                return state;
            }
        }

        // save value to value state
        valueState.value.control.push(value);
        valueState.value.sp = JSON.stringify(valueState.value.control);
        valueState.errorMessage = "";
        this.setValueState_for(internalName, valueState);

        // return new textbox state
        return state;
    }

    /** Array data Delete item event handler
     * @param internalName Field InternalName used as the fields reference ID
     * @param ndx Index of item that was clicked
     */
    private arrayData_onDelete(internalName: string, ndx: number) {
        CustomDialog.confirm("Are you sure you want to delete this item?").then(response => {
            if (response === true) {
                this.updateValueState_ArrayData(internalName, null, ndx, true);
            }
        });
    }

    /** Validated / ValidatedBy display with button
     * @param internalName Date field InternalName used to store the validated date
     * @param validatedByFieldName PeoplePicker field InternalName used to store the validated by user info
     */
    private async validated_onClick(internalName: string, validatedByFieldName: string) {
        // ensure user
        if (!this.user) this.user = await this.props.list.getCurrentUser();
        
        const today = new Date(),
            valueState = this.getValueState_by_InternalName(internalName),
            validatedByState = this.getValueState_by_InternalName(validatedByFieldName);

        // add date & user field to modified list
        this.fieldWasModified(internalName);
        this.fieldWasModified(validatedByFieldName);

        // update value state for both fields
        valueState.value = today.toISOString();
        validatedByState.value.control = this.convertSiteUser_toUserInfo(this.user);
        validatedByState.value.sp = this.user.Id;

        // set new value states
        this.setValueState_for(internalName, valueState);
        this.setValueState_for(validatedByFieldName, validatedByState);
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
            fieldOption = this.getFieldWebPartOptions_by_InternalName(internalName),
            isRequired = field.Required || fieldOption.required || false;

        let error: string = null,
            tempValue;

        switch (field["odata.type"]) {
            case "SP.FieldText":
                error = this.validateString(value, isRequired, field.MaxLength || 255);
                break;

            case "SP.FieldNumber":
                error = this.validateNumber(value, isRequired, field.MinimumValue || null, field.MaximumValue || null);
                break;

            case "SP.FieldMultiLineText":
                error =  this.validateString(this.getTextFromHTMLString(value), isRequired);
                break;
        
            case "SP.FieldChoice":
                error = this.validateChoice(value, isRequired);
                break;
                    
            case "SP.FieldMultiChoice":
                tempValue = value && value.control ? value.control : value;
                error = this.validateChoice(tempValue, isRequired);
                break;

            case "SP.FieldLookup":
                error  = this.validateChoice(value.control, isRequired);
                break;

            case "SP.FieldUrl":
                tempValue = value && value.Url ? value.Url : value;
                error = this.validateString(tempValue, isRequired, 255);
                break;

            case "SP.FieldDateTime":
                error = this.validateDate(value, isRequired);
                break;

            case "SP.FieldUser":
                error = this.validateChoice(value.control, isRequired);
                break;

            default:
                LOG("Unhandled field validation:", field["odata.type"], field.InternalName, value);
                break;
        }

        // validate value based on field

        return error;
    }

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
     * @param choices Array of values or singlar value
     * @param isRequired Boolean if the field is required/mandatory
     * @param min Minimum number of selections
     * @param max Maximum number of selections
     * @returns Error message or null
     */
    private validateChoice(choices: any | any[], isRequired: boolean = false, min?: number, max?: number): string {
        if (typeof choices === "string") return this.validateString(choices, isRequired);
        if (typeof choices === "number") return this.validateNumber(choices, isRequired);

        const length = choices instanceof Array ? choices.length : 0;

        if (isRequired && length === 0) return "Required.";

        if (typeof min === "number" && length < min) return `Not enough. Select a miniumum of ${min} choice${min !== 1 ? "" : "s"}.`;
        
        if (typeof max === "number" && length > max) return `Too much. Select a maximum of ${max} choice${max !== 1 ? "" : "s"}.`;

        return null;
    }
//#endregion
    
//#region SAVE, CANCEL, DELETE
    /** New or Edit form command bar
     * @returns DIV with 'Cancel', 'Save command bar
     */
    private NewEditCommandBar(): React.ReactElement {
        const buttons = [],
            stackStyle: IStackStyles = {
                root: { height: 44 }
            },
            tokens: IStackTokens = { childrenGap: 10 };

        // save
        if (this.state.canSaveForm) {
            buttons.push(React.createElement(CommandBarButton, {
                hidden:  !this.state.canSaveForm,
                disabled: this.state.isProcessingForm,
                iconProps: this.props.form === "new" ? { iconName: "Add" } : { iconName: "Save" },
                text: this.props.form === "new" ? "Create" : "Update",
                className: styles.button + " " + styles.buttonPrimary,
                onClick: this.save_onClick.bind(this)
            }));
        }

        // cancel
        if (this.state.canCancelForm) {
            buttons.push(React.createElement(CommandBarButton, {
                hidden: !this.state.canCancelForm,
                disabled: this.state.isProcessingForm,
                iconProps: { iconName: "Cancel" },
                text: "Cancel",
                className: styles.button,
                onClick: this.cancel_onClick.bind(this)
            }));
        }

        // delete
        if (this.props.form === "edit" && this.state.canDeleteItem) {
            buttons.push(React.createElement(CommandBarButton, {
                hidden:  !this.state.canSaveForm,
                disabled: this.state.isProcessingForm,
                iconProps: { iconName: "Delete" },
                text: "Delete",
                className: styles.button + " " + styles.buttonPrimary,
                onClick: this.delete_onClick.bind(this)
            }));
        }

        return <Stack 
                horizontal
                horizontalAlign="end"
                className={styles.commandbar}
                styles={stackStyle}
                tokens={tokens}
            >{ buttons }</Stack>;
    }

    /** Cancel button click handler
     * @param evt Click event object
     * @param item Toolbar item clicked
     */
    private async cancel_onClick(
        evt?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>,
        item?: IContextualMenuItem): Promise<boolean | void> {

        let goBack: boolean;

        // check to see if any changes were made to the form
        if (this.fieldsThatHaveBeenModified.length > 0) {
            // make sure they didn't accidentally clicked 'Cancel'
            goBack = await CustomDialog.confirm(`Are you sure you want to cancel? ${(this.props.form === "new") ? 
                "The new item" : "Changes"} will not be saved.`, "Cancel?", undefined);

        } else {
            // no changes, ok to cancel
            goBack = true;
        }
        
        if (goBack) this.goBack();
    }

    /** Save button click handler
     * @param evt Click event object
     * @param item Toolbar item clicked
     */
    private async save_onClick(
        evt?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>,
        item?: IContextualMenuItem): Promise<boolean | void> {

        this.setState({"isProcessingForm": true});

        // show modal

        const _getSPValueFromValueState = valueState => {
            if (valueState.value !==  null && typeof valueState.value === "object" && "sp" in valueState.value) {
                return valueState.value.sp;
            } else {
                return valueState.value;
            }
        };

        const _getSPFieldName = internalName => {
            const field = this.getField_by_InternalName(internalName);
            if (field === null) return internalName;

            switch (field["odata.type"]) {
                case "SP.FieldLookup":
                case "SP.FieldUser":
                    return internalName + "Id";
            
                default:
                    return internalName;
            }
        };

        const _abort = msg => {
            this.setState({"errorMessage": msg, "isProcessingForm": false});
            return false;
        };
        
        // build the payload (form data)
        const data = {};
        if (this.props.form === "new") {
            // new forms: save fields that have changed and items that have values (default values)
            for (let i = 0; i < this.props.list.fields.length; i++) {
                const internalName = this.props.list.fields[i].InternalName,
                    valueState = this.getValueState_by_InternalName(internalName);

                // check if valuestate exists.
                if (valueState === null) continue;

                const field = this.getField_by_InternalName(internalName),
                    value = _getSPValueFromValueState(valueState),
                    wasModified = this.fieldsThatHaveBeenModified.indexOf(internalName) > -1,
                    errorMessage = this.validateFieldValue(internalName, valueState.value);

                if (errorMessage !== valueState.errorMessage) {
                    valueState.errorMessage = errorMessage;
                    this.setValueState_for(internalName, valueState);
                }
                if (errorMessage) return _abort(`ERROR! ${field.Title}: ${valueState.errorMessage}`);

                if (value !== null || wasModified) {
                    data[_getSPFieldName(internalName)] = value;
                }
            }

        } else {
            // edit form: update fields that were changed
            this.fieldsThatHaveBeenModified.forEach(internalName => {
                const valueState = this.getValueState_by_InternalName(internalName),
                    field = this.getField_by_InternalName(internalName),
                    value = _getSPValueFromValueState(valueState),
                    errorMessage = this.validateFieldValue(internalName, valueState.value);

                if (errorMessage !== valueState.errorMessage) {
                    valueState.errorMessage = errorMessage;
                    this.setValueState_for(internalName, valueState);
                }
                if (errorMessage) return _abort(`ERROR! ${field.Title}: ${valueState.errorMessage}`);

                data[_getSPFieldName(internalName)] = value;
            });
        }

        // submit the data
        if (Object.keys(data).length > 0) {
            let response: IItemAddResult;

            this.setState({"errorMessage": null});
            
            try {
                if (this.props.form === "new") {
                    // new = add item
                    response = await this.props.list.api.items.add(data);
                    this.listItem = response.data;
                } else {
                    // edit = update item
                    response = await this.props.list.api.items.getById(parseInt(this.listItem.Id)).update(data);
                }

                // if user can update permissions (full control), update item permissions
                const userCanUpdate = await this.props.list.api.currentUserHasPermissions(PermissionKind.ManagePermissions);
                if (userCanUpdate) await this.props.list.updateContentManagers(this.listItem.Id);

                // if successful, go back
                return this.goBack();

            } catch(er) {
                LOG("ERROR! Unable to save item:", er);
                this.setState({
                    "isProcessingForm": false,
                    "errorMessage": `Unable to ${ this.props.form === "new" ? "create" : "update" } item. See console for more details.`
                });
            }
        } else {
            // if on edit form and user clicked 'save' and no data needed to be updated, then just go back
            if (this.props.form === "edit") {
                return this.goBack();
            }
        }

    }

    /** Delete button click handler
     * @param evt Click event object
     * @param item Toolbar item clicked
     */
    private async delete_onClick(
        evt?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>,
        item?: IContextualMenuItem): Promise<boolean | void> {

        if (this.listItem && this.listItem.Id) {
            this.setState({"isProcessingForm": true});

            const comfirmed = await CustomDialog.confirm(
                `Are you sure want to delete '${ this.listItem.Title } (ID: ${ this.listItem.Id })'?`,
                "Confirm Delete", {yes: "Delete", no: "Cancel"});
            if (comfirmed) {
                const modalMsg = CustomDialog.modalMsg("Processing...", "Please wait!");

                try {
                    await this.props.list.api.items.getById(this.listItem.Id).delete();
                    modalMsg.close();
                    return this.goBack();

                } catch (er) {
                    LOG(`ERROR! Unable to delete item '${ this.listItem.Title } (ID: ${ this.listItem.Id })'`, er);
                    modalMsg.close();
                    this.setState({
                        "isProcessingForm": false,
                        "errorMessage": "Unable to delete item. See console for more details."
                    });
                }

            }

        }
    }

    private goBack(): void {
        const url = new URL(location.href);
        this.listItem = null;

        if (this.props.history.length > 1) {
            window.history.back();
        } else {
            url.searchParams.delete(`${this.props.properties.urlParam}form`);
            location.assign(url.toString());
        }
    }

//#endregion

//#region VALUESTATE: LIST ITEM VALUE STATE
    /** Get the ValueState ({defaultValue, value, errorMessage}) for this field
     * @param internalName Field InternalName for the ValueState object to retrieve
     * @returns ValueState object for Field ID passed
     */
    public getValueState_by_InternalName(internalName: string): IAboutUsValueState {
        const key = this.valueStateKeyPrefix + internalName;
        return (key in this.state) 
            ? this.state[key] 
            : null ;
    }

    /** Initialize a new ValueState: {defaultValue, value, errorMessage}
     * @param internalName Field InternalName
     * @param value Initial or default value
     * @returns ValueState object
     */
    private initializeFieldValueState(internalName: string, value?: any): IAboutUsValueState {
        const wpOption = this.getFieldWebPartOptions_by_InternalName(internalName);
        if (value === undefined) value = null;
    
        return {
            defaultValue: value,
            value: value,
            errorMessage: null,
            disabled: (wpOption.controlled && !this.state.isAdmin)
        };
    }

    /** Get the ValueState for a field. If it doesn't exists, initializes a new ValueState and add to state.
     * @param internalName Field InternalName for the ValueState object to create or retrieve
     * @param defaultValue Default value if/when the ValueState is initialized. Does not update the existing ValueState.
     * @returns ValueState object
     */
    public ensureValueState_for(internalName: string, defaultValue?: any): IAboutUsValueState {
        const field = this.getField_by_InternalName(internalName);
        let valueState = this.getValueState_by_InternalName(field.InternalName);

        // if null, create the value state
        if (valueState === null) {
            valueState = this.initializeFieldValueState(field.InternalName, defaultValue || null);

            // keep track of newly created value states
            const key = this.valueStateKeyPrefix + internalName;
            this.state = {...this.state, [key]: valueState};
        }

        return valueState;
    }

    /** Update the ValueState for a specific field. Updates the state and display. Pauses, to allow the state to update correctly.
     * @param internalName Field InternalName for the ValueState object to set.
     * @param valueState New ValueState to set.
     */
    public async setValueState_for(internalName: string, valueState: IAboutUsValueState): Promise<void> {
        const key = this.valueStateKeyPrefix + internalName;
        this.setState({[key]: valueState});
        // setState needs just a momemt to update the state. 
        // not required but helps if you need to reference it right away.
        await sleep(0);
    }

    /** Creates the MultiChoice object (IAboutUsMultiChoiceItemValue)
     * @param choices Choice value or array of values.
     * @returns Multi-Choice value state. 
     */
     private createValueState_MultiChoice(choices: string | string[]): IAboutUsMultiChoiceItemValue {
        // if a string was passed, assume it is an SP delimited (;#) string
        if (typeof choices === "string") choices = this.parseSPDelimitedStringValues(choices);

        // exit if already an IAboutUsMultiChoiceItemValue
        if (choices && typeof choices === "object" && "sp" in choices) return choices;

        const value = {
            sp: { results: [] },
            control: null
        };

        if (choices instanceof Array) {
            value.sp.results = choices;
            value.control = choices;
        }

        return value;
    }

    /** Creates the LookupData object (IAboutUsLookupItemValue)
     * @param choices Lookup value or array of lookup values. Object(s) must contain 'ID' property
     * @param internalName Field InternalName used as the field's reference ID.
     * @returns LookupValue object that can be used for the Combo
     */
    private createValueState_LookupItem(choices: any | any[], internalName: string): IAboutUsMultiChoiceItemValue {
        // if "choices" is the valueState
        if (choices && typeof choices === "object" && "sp" in choices && "control" in choices) return choices;
        
        const field = this.getField_by_InternalName(internalName),
            lookupField = field.LookupField || "Id",
            value: IAboutUsMultiChoiceItemValue = {
                sp: field.AllowMultipleValues ? {results: []} : null,
                control: []
            };

        // return early if no value(s)
        if (!choices) return value;

        // normalize values. make into array
        if (choices instanceof Array === false) choices = [choices];

        // loop through each value
        for (let i = 0; i < choices.length; i++) {
            const selected = choices[i],
                // get the selected value key
                key = selected.Id || selected.ID || selected.id || selected.key || null,
                name = ("name" in selected) ? selected.name : (lookupField in selected) ? selected[lookupField] : null;
                
            const control = {"key": key, "name": name};

            if (key) {
                if (field.AllowMultipleValues) {
                    value.sp.results.push(key);
                } else {
                    value.sp = key;
                }
                value.control.push(control);

            }
        }

        return value;
    }

    /** Creates the UserData object (IAboutUsUserValue).
     * Types of user data accepted:
     * - SP.FieldUser (single) = IUserInfoItem = {ID: number, Name: string}
     * - SP.FieldUser (multi) = IUserInfoItem[] = {results: [{ ID: number, Name: string }, ...]}
     * - PeoplePicker (SPFX PnP React Control) onChange = IPersonaProps[] = 
     *      [{id: number, loginName: string, secondaryText: string}, ...]
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
        // check to see if users is IAboutUsUserValue object
        if (users && typeof users === "object" && "sp" in users && "control" in users) return users;

        const field = this.getField_by_InternalName(internalName),
            value: {sp: any | any[], control: any | any[]} = {
                sp: (field.AllowMultipleValues) ? {results: []} : -1,
                control: []
            };

        const _getControlValue = user => {
            return (
                // IPersonaProps
                user.loginName ||
                user.secondaryText ||
                // ISiteUserProps
                user.LoginName ||
                user.Email ||
                // IUserInfoItem
                user.Name ||
                // ValueState (recalled, passing the valuestate)
                users.control ||
                // none of those properties exists
                ""
            ).split("|").pop();
        };
        const _getIdValue = user => {
            return (
                // IPersonaProps
                user.id ||
                // ISiteUserProps
                user.Id ||
                // IUserInfoItem
                user.ID ||
                // ValueState (recalled, passing the valuestate)
                users.sp ||
                // none of those properties exists
                null
            );
        };

        // return early if users data is null
        if (!users) return value;

        // normalize users data. make all data an array of user infor (SPUSer, IUserInfoItem, IPersona)
        if (users instanceof Array === false) users = [users];

        //  loop through each user info
        for (let i = 0; i < users.length; i++) {
            const user = users[i],
                controlValue = _getControlValue(user),
                idValue = _getIdValue(user);

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

    /** Creates the ComplesData object (IAboutUsComplexValue)
     * @param data Stringified representation of the complex data or the complex data type itself
     * @returns ValueState value with values for the control and SP API.
     */
    private createValueState_ComplexData(data: string | TAboutUsComplexData | IAboutUsComplexValue): IAboutUsComplexValue {
        const value: IAboutUsComplexValue = {
            "control": [],
            "sp": ""
        };

        // exit early if complex data array was passed
        if (data && (typeof data === "object") && "conrol" in data && "sp" in data) return data as IAboutUsComplexValue;

        // exit early if null or empty
        if (!data || (data as string | TAboutUsComplexData).length === 0) return value;

        // parse the data based on the type passed
        if (typeof data === "string") {
            data = trim(data);

            if ((/(^\[).*(\]$)/gi).test(data)) {
                // data is stringified array
                data = JSON.parse(data);
            }
        }
        
        if (data instanceof Array) {
            // remove empty items
            data.forEach(item => {
                if (item && (typeof item === "object") && Object.keys(item).length > 0) value.control.push(item);
            });
        } 

        value.sp = (value.control.length > 0) ? JSON.stringify(value.control) : "";

        return value;
    }

    /** Creates the string array value with SP stringified value.
     * @param data Stringified representation of the string array or the 'Enterprise Keyword' string list.
     * @returns ValueState value with values for the control and SP API.
     */
    private createValueState_Keywords(data: string | string[] | IAboutUsKeywordsValue): IAboutUsKeywordsValue {
        const value: IAboutUsKeywordsValue = {
            "control": [],
            "sp": ""
        };

        // exit early if keyword valut state was passed
        if (data && (typeof data === "object") && "conrol" in data && "sp" in data) return data;

        // exit early if null or empty
        if (!data || (data as string | string[]).length === 0) return value;

        // parse data based on the type passed
        if (typeof data === "string") {
            data = trim(data);

            // what type of string?
            if ((/(^\[).*(\]$)/gi).test(data)) {
                // data is a stringified array
                data = JSON.parse(data);

            } else {
                // data is a list of words. possibly list of Enterprise Keywords
                data = data.split(";").map(keyword => keyword.split("|")[0]);

            }
        } 
        
        if (data instanceof Array) {
            // remove empty items
            data.forEach(keyword => {
                if (typeof keyword === "string") {
                    keyword = trim(keyword);
                    if (keyword.length > 0) value.control.push(keyword);
                }
            });
        }
        
        value.sp = (value.control.length > 0) ? JSON.stringify(value.control) : "";

        return value;
    }

    /** Add, update, or delete complex data item.
     * @param internalName Field InternalName used as the field's reference ID
     * @param value New or updated value. Pass null, if deleting the item.
     * @param ndx Index of the item to update or delete.
     * @param deleteItem Flag to delete item. Note: ndx must be valid.
     */
    private updateValueState_ArrayData(internalName: string, value?: any, ndx?: number, deleteItem?: boolean) {
        const valueState = this.getValueState_by_InternalName(internalName);

        if (!("control" in valueState.value)) valueState.value = {"sp": null, "control": []};
        if (!(valueState.value.control instanceof Array)) valueState.value.control = [];

        if (typeof ndx === "number" && ndx > -1 && ndx < valueState.value.control.length) {
            // valid index. update or delete?

            if (deleteItem) {
                // delete the item
                valueState.value.control.splice(ndx, 1);

            } else if (value) {
                // update the item
                valueState.value.control[ndx] = value;
            }

        } else if (value) {
            // add the item
            valueState.value.control.push(value);
        }

        // update SP API value
        valueState.value.sp = JSON.stringify(valueState.value.control);

        // add field to modifed list
        this.fieldWasModified(internalName);

        // trigger update
        this.setValueState_for(internalName, valueState);
    }

    /** Rearranges the order of array objects
     * @param internalName Field InternalName used as the field's reference ID
     * @param oldIndex Old index of the item.
     * @param newIndex New index for the item that moved.
     */
     private reorderValueState_ArrayData(internalName: string, oldIndex: number, newIndex: number) {
        const valueState = this.getValueState_by_InternalName(internalName);

        if (!("control" in valueState.value)) valueState.value = {"sp": null, "control": []};
        if (!(valueState.value.control instanceof Array)) valueState.value.control = [];

        const value = valueState.value.control;

        if (value && value instanceof Array) {
            // update valueState with new value
            valueState.value.control = AboutUsDisplay.rearrangeArray(value, oldIndex, newIndex);
            valueState.value.sp = JSON.stringify(valueState.value.control);

            // add field to modifed list
            this.fieldWasModified(internalName);
        
            // trigger update
            this.setValueState_for(internalName, valueState);
        }    
    }
//#endregion

//#region HELPERS
    /** Creates the Broadcast button for each of the complex data items.
     * @param ndx Index of the complex data item
     * @param value Value of the complex data item
     * @returns Command Bar Items array
     */
    private commandbarBroadcastButton(ndx: number, value: TAboutUsComplexData): ICommandBarItemProps[] {
        // broadcast button
        const extraButtons = [],
            broadcastDate = (value.broadcastDate) ? new Date(value.broadcastDate) : null,
            broadcasting = isInRange_numDays(broadcastDate, this.props.properties.broadcastDays);
        
        extraButtons.push({
            key: `btnBroadcast${ndx}`,
            text: (broadcasting) ? "Stop broadcasting" : "Start broadcasting",
            iconProps: { iconName: (broadcasting) ? "UserRemove" : "NetworkTower" },
            iconOnly: true,
            ariaLabel: (broadcasting) ? "Stop broadcasting bio" : "Start broadcasting bio",
            buttonStyles: { root: { "border": "1px solid", "border-radius": "3px;" }},
            onClick: ()=>{ this.updateBroadcastDate("Bios", ndx, (broadcasting) ? null : new Date()); }
        });

        return extraButtons;
    }

    /** Updates the BroadcastDate field with the most current date and updates the complex data field.
     * @param internalName Field InternalName used as the field's reference ID
     * @param ndx Index for the complex data item
     * @param newBroadcasatDate Updated broadcast date
     */
    public updateBroadcastDate(internalName: string, ndx: number, newBroadcasatDate: Date) {
        const valueState = this.getValueState_by_InternalName(internalName),
            broadcastDateFieldName = "BroadcastDate",
            broadcastDateValueState = this.getValueState_by_InternalName(broadcastDateFieldName),
            startingDate = new Date(0);

        // make sure BroadcastDate field is created.
        if (!broadcastDateValueState) {
            LOG("ERROR! 'BroadcastDate' date field is missing. Could not update broadcast values.");
            return;
        }

        let mostCurrentDate = new Date(startingDate.toISOString());

        // loop through each item's broadcast date. get the most current date
        valueState.value.control.forEach((item, i) => {
            // update item if matching
            if (i === ndx) item.broadcastDate = (newBroadcasatDate) ? newBroadcasatDate.toISOString() : "";
            const itemBroadcastDateValue = item.broadcastDate,
                itemBroadcastDate = (itemBroadcastDateValue) ? new Date(itemBroadcastDateValue) : null;

            if (itemBroadcastDate && itemBroadcastDate > mostCurrentDate) {
                mostCurrentDate = new Date(itemBroadcastDateValue);
            }
        });

        // set BroadcastDate field value
        broadcastDateValueState.value = (mostCurrentDate > startingDate) ? mostCurrentDate.toISOString() : null;

        // update BroadcastDate value state
        this.fieldWasModified(broadcastDateFieldName);
        this.setValueState_for(broadcastDateFieldName, broadcastDateValueState);

        // reset bio field 'sp' value
        valueState.value.sp = JSON.stringify(valueState.value.control);

        // update bio value state
        this.fieldWasModified(internalName);
        this.setValueState_for(internalName, valueState);
    }

    /** Generates a new MicroForm based on the InternalName passed.  
     * @param internalName Field InternalName used as the fields's reference ID.
     * @param formValues Default form values. If null, form will be populated with field default values.
     * @returns New MicroForm instance for internal name passed.  
     *      (async) .show() returns the form values as JSON or null if cancelled.
     */
    private generateMicroForm(internalName: string, formValues?: Record<string, any>): AboutUsMicroForm {
        const field = this.getField_by_InternalName(internalName),
            titleValue = this.getValueState_by_InternalName("Title"),
            template = (internalName in DataFactory.listTemplate.fieldMicroForms) ? 
                DataFactory.listTemplate.fieldMicroForms[internalName] : null,
            microForm = (template)
                ? new AboutUsMicroForm(
                    this.props.ctx,
                    this.props.properties,
                    `${field.Title} Form`,
                    template, formValues,
                    null,
                    trim(titleValue.value)
                )
                : null;

        return microForm;
    }

    /** Get list item value for a specif field. Assumes the list item's data was already retrieved.
     * @param internalName Field InternalName used as the fields's reference ID.
     * @returns List item value for the field
     */
    private getListItemData_for(internalName: string): any {
        return (internalName in this.listItem) ? this.listItem[internalName] : null ;
    }

    /** Get Field information
     * @param internalName Field InternalName of field to retrieve
     * @returns Field info
     */
     private getField_by_InternalName(internalName: string): IDataFactoryFieldInfo {
        return find(this.props.list.fields, ["InternalName", internalName]) || null;
    }

    /** get field options from web part property pane
     * @param internalName Field InternalName of field to retrieve
     * @returns Field option from web part property pane
     */
    private getFieldWebPartOptions_by_InternalName(internalName: string): IAboutUsAppFieldOption {
        const field = this.getField_by_InternalName(internalName);
        let fieldOption: IAboutUsAppFieldOption = {
            required: field.Required || false,  // default value
            controlled: false                   // default value
        };

        if (internalName in this.props.properties.fields) fieldOption = this.props.properties.fields[internalName];

        return fieldOption;
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
        if (typeof delimitedString === "string" && delimitedString.indexOf(";#") > -1) 
            return trim(delimitedString).replace(/^(;#)|(;#)$/g,"").split(";#");

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

    /** Get the text from a string representing HTML.
     * @param htmlString String representing HTML.
     * @returns innerText from the HTML
     */
    private getTextFromHTMLString(htmlString: string): string {
        this._htmlNode.innerHTML = htmlString || "";

        return this._htmlNode.innerText;
    }

    /** Converts the site user object to the IUserInfo object.
     * ISiteUserInfo object is the user object returned by the PNP/SP Site-User REST request
     * AboutUsDisplay.IUserInfo object is the expanded User data from the Items REST request.
     * @param user Site user information object
     */
    private convertSiteUser_toUserInfo(user: ISiteUserInfo): AboutUsDisplay.IUserInfo {
        return (user) ? {
            "odata.type": user["odata.type"],
            "odata.id": user["odata.id"],
            "ID": user.Id,
            "Title": user.Title,
            "Name": user.LoginName,
            "EMail": user.Email
        } : null;
    }

    /** Flag (or unflag) fields that have been modified.
     * @param internalName Field InternalName used as the field's reference ID.
     * @param addToList Add or remove from fieldsThatHaveBeenModified list. Default: true;
     */
    private fieldWasModified(internalName: string, addToList: boolean = true) {
        const index = this.fieldsThatHaveBeenModified.indexOf(internalName);
        if (addToList && index === -1) {
            // add to list
            this.fieldsThatHaveBeenModified.push(internalName);

        } else if (addToList === false && index > -1) {
            // remove from list
            this.fieldsThatHaveBeenModified.splice(index, 1);
        }
    }

    /** sets user permission flag states (add, edit, delete) */
    private async setCurrentUserFlags(): Promise<void> {
        const userPermissions = await this.props.list.getUserPermissions();

        this.setState({
            canSaveForm: userPermissions.canAdd || userPermissions.canEdit,
            canDeleteItem: userPermissions.canDelete,
            isAdmin: userPermissions.canAdd && userPermissions.canDelete
        });
    }

//#endregion
}