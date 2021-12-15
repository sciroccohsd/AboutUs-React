// React and MS Fabric UI controls for form fields
import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { assign, result } from 'lodash';
// https://docs.microsoft.com/en-us/javascript/api/office-ui-fabric-react?view=office-ui-fabric-react-latest
import { Dropdown,
    IDropdownProps,
    ILabelProps,
    ISpinnerProps,
    ITextFieldProps, 
    Label,
    ActionButton,
    Spinner, 
    TextField, 
    IButtonProps,
    IconButton,
    TooltipHost} from 'office-ui-fabric-react';

// https://pnp.github.io/sp-dev-fx-controls-react/index.html
//> npm install @pnp/spfx-controls-react --save
import * as ReactControls from '@pnp/spfx-controls-react';
import { ComboBoxListItemPicker, 
    DateTimePicker, 
    FilePicker, 
    IComboBoxListItemPickerProps, 
    IDateTimePickerProps, 
    IFilePickerProps, 
    IFilePickerResult, 
    IListItemPickerProps, 
    IPlaceholderProps, 
    ListItemPicker } from '@pnp/spfx-controls-react';
import { IRichTextProps, RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { PeoplePicker, IPeoplePickerProps } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import DataFactory, { IFieldUrlValue } from './DataFactory';
import * as AboutUsDisplay from './AboutUsDisplay';
import { arrayListFormat, DEBUG, DEBUG_NOTRACE, IAboutUsAppWebPartProps, LOG } from '../AboutUsAppWebPart';
import CustomDialog from './CustomDialog';
import { IItem } from '@pnp/sp/items';

//#region EXPORTS
export default ReactControls;
export * from '@pnp/spfx-controls-react/lib/RichText';
export * from '@pnp/spfx-controls-react/lib/PeoplePicker';
export * from '@pnp/spfx-controls-react/lib/Toolbar';
//#endregion


//#region WEB PART PLACEHOLDER: "Configure your web part"
export class ShowConfigureWebPart extends React.Component<Partial<IPlaceholderProps>> {
    public render(): React.ReactElement<Partial<IPlaceholderProps>> {
        return <ReactControls.Placeholder
            description={ this.props.description }
            iconName={ this.props.iconName || "Edit" }
            buttonLabel={ this.props.buttonLabel || "Configure" }
            iconText={ this.props.iconText || "Configure your web part" }
            onConfigure={ this.props.onConfigure }
        />;
    }
}
//#endregion

//#region LOADING...
export class LoadingSpinner extends React.Component<ISpinnerProps> {
    public render(): React.ReactElement<ISpinnerProps> {
        const props = assign({
                label: "Loading...",
                ariaLive: "assertive",
                labelPosition: "right"
            }, this.props);

        return (
            <div>
                { React.createElement(Spinner, props) }
            </div>
        );
    }
}
//#endregion

//#region BASIC CONTROLS:  LABEL, ERROR, DESCRIPTION
export interface IFieldWrapperProps extends React.HTMLAttributes<HTMLDivElement> {
    required?: boolean;
}
export class FieldWrapper extends React.Component<IFieldWrapperProps> {
    public render(): React.ReactElement<React.HTMLAttributes<HTMLDivElement>> {
        const className = [ this.props.className || "", styles.fieldWrapper];

        const props = assign({}, this.props, { className: className.join(" ") });

        return React.createElement("div", props);
    }
}

export class LabelElement extends React.Component<ILabelProps> {
    public render(): React.ReactElement<ILabelProps> {
        const props = assign({}, this.props, { className: "FormControlsLabel "  + (this.props.className || "")});
        return React.createElement(Label, props);
    }
}

export interface IDescriptionProps {
    text?: string;
}
export class DescriptionElement extends React.Component<IDescriptionProps> {
    public render(): React.ReactElement<IDescriptionProps> {
        return (
            <>
                { this.props.text ? 
                    <div className="FormControlsDescription">
                        <span className={ `ms-TextField-description ${ styles.description }`}>
                            { this.props.text }
                        </span>
                    </div>
                    : null
                }
            </>
        );
    }
}

export interface IErrorMessageProps {
    text?: string;
}
export class ErrorMessageElement extends React.Component<IErrorMessageProps> {
    public render(): React.ReactElement<IErrorMessageProps> {
        return (
            <>
                { this.props.text ? 
                    <div role="alert">
                        <p className={`ms-TextField-errorMessage ${styles.errorMessage}`}>
                            <span data-automation-id="error-message" >{ this.props.text }</span>
                        </p>
                    </div>
                    : null
                }
            </>
        );
    }
}
//#endregion

//#region TEXTBOX CONTROL
export class TextboxControl extends React.Component<ITextFieldProps> {
    public render(): React.ReactElement<ITextFieldProps> {
        return (
            <FieldWrapper>
                { React.createElement(TextField, this.props) }
            </FieldWrapper>
        );
    }
}
//#endregion

//#region URL CONTROL
export interface IUrlControlProps {
    id: string;
    className?: string;
    required?: boolean;
    disabled?: boolean;
    label?: string;
    description?: string;
    defaultValue?: IFieldUrlValue;
    showTextField?: boolean;
    errorMessage?: string;
    onChange?: {(url: string, text: string, id: string):void};

    externalRepo?: string;
    folderName?: string;      
    filePickerProps?: Partial<IFilePickerProps>;
}
export interface IUrlControlState {
    url: string;
    text: string;
    defaultFolderAbsolutePath?: string[];
}

export class UrlControl extends React.Component<IUrlControlProps, IUrlControlState> {
    constructor(props) {
        super(props);

        const urlValue = this.props.defaultValue ? this.props.defaultValue.Url : null,
            textValue = this.props.defaultValue ? this.props.defaultValue.Description : urlValue,
            defaultFolderAbsolutePath = [this.props.filePickerProps.context.pageContext.web.absoluteUrl];
        
        // add external repo as the file picker default location
        if (this.props.filePickerProps && this.props.externalRepo) {
            defaultFolderAbsolutePath.push(this.props.externalRepo.split("/").pop());
        }


        this.state = {
            url: urlValue,
            text: textValue,
            defaultFolderAbsolutePath: defaultFolderAbsolutePath
        };
    }
    public render(): React.ReactElement<IUrlControlProps> {
        const showTextField = typeof this.props.showTextField === "boolean" ? this.props.showTextField : true ,
            onChange = this.props.onChange ? evt => {
                const isUrl = evt.target.id === this.props.id,
                    key: string = (isUrl) ? "url" : "text",
                    value = evt.target.value;
                    
                this.setState({...this.state, [key]: value});

                this.props.onChange(value, key, this.props.id);
            } : null,
            urlProps: ITextFieldProps = {
                type: "url",
                id: this.props.id,
                className: `FormControlsUrlValue ${this.props.className || ""}`,
                label:  this.props.label,
                value: this.state.url,
                required: this.props.required,
                disabled: this.props.disabled,
                placeholder: "https://...",
                
                onChange: onChange
            },
            textProps: ITextFieldProps = {},
            filePickerProps: IFilePickerProps = (this.props.filePickerProps)
                ? assign({
                    defaultFolderAbsolutePath: this.state.defaultFolderAbsolutePath.join("/"),
                    onSave: this.filePicker_onSave.bind(this), // need to handle uploads
                    onChange: this.filePicker_onChange.bind(this),
                    buttonLabel: "Select a File",
                    buttonIcon: "FileImage",
                }, this.props.filePickerProps) as IFilePickerProps
                : null;

        if (showTextField) {
            // show both URL and Text fields, populate properties for text field
            textProps.disabled = this.props.disabled;
            textProps.id = this.props.id + "_text";
            textProps.className = `FormControlsTextValue ${this.props.className || ""}`;
            textProps.label = `Text for: ${ this.props.label || "Url"}`;
            textProps.value = this.state.text;
            textProps.description = this.props.description;
            textProps.placeholder = "(Optional) Alternative text";
            textProps.errorMessage = this.props.errorMessage;
            textProps.onChange = onChange;
        } else {
            // show only the URL field, add additional properties to the URL field
            urlProps.description = this.props.description;
            urlProps.errorMessage = this.props.errorMessage;
        }

        return (
            <FieldWrapper>
                { React.createElement(TextField, urlProps) }
                { showTextField ? React.createElement(TextField, textProps) : null }
                { filePickerProps ? React.createElement(FilePicker, filePickerProps) : null }
           </FieldWrapper>
        );
    }

    public async componentDidMount() {
        // ensure external repo folder exists
        if (this.props.filePickerProps && this.props.externalRepo && this.props.folderName) {
            const exists = await DataFactory.ensureFolder(this.props.externalRepo, this.props.folderName);
            if (exists) {
                const defaultFolderAbsolutePath = [...this.state.defaultFolderAbsolutePath];
                defaultFolderAbsolutePath.push(this.props.folderName);
                this.setState({"defaultFolderAbsolutePath": defaultFolderAbsolutePath});
            }
        }
    }

    //#region FILE PICKER
    private updateFields(item: IFilePickerResult) {
        const url = item.fileAbsoluteUrl,
            text = item.fileName;

        // if no url, item  may need to be uploaded.
        if (!url) return;

        this.setState({
            url: url,
            text: text
        });

        this.props.onChange(url, "url", this.props.id);
        this.props.onChange(text, "text", this.props.id);
    }

    private async uploadItem(item: IFilePickerResult): Promise<IItem> {
        const file = await item.downloadFileContent();
        let uploadedItem: IItem = null;

        // invalid file type
        if (this.props.filePickerProps.accepts && this.props.filePickerProps.accepts.length > 0) {
            if (this.props.filePickerProps.accepts.indexOf(file.name.toLowerCase().split(".").pop()) === -1) {
                // not accepted file type
                await CustomDialog.alert(
                    `Unable to upload file. Allowed types: ${arrayListFormat(this.props.filePickerProps.accepts, ", ", " or ")}`,
                    "Invalid file type");

                return uploadedItem;
            }
        }

        // only upload files that are 5MBs or lower
        if (file.size <= 5500000) {
            uploadedItem = await DataFactory.uploadFile(this.props.externalRepo, file, this.props.folderName);
                
        } else {
            // too large
            await CustomDialog.alert("Unable to upload file. Image size must be less than 5 MBs.", "File too large");
        }

        return uploadedItem;
    }

    private filePicker_onChange(items: IFilePickerResult[]) {
        // the URL field can only accept one item
        const item = (items && items.length > 0) ? items[0] : null;
        if (item) this.updateFields(item);
    }

    private async filePicker_onSave(items: IFilePickerResult[]) {
        // external repo not set up
        if (!this.props.externalRepo) {
            return await CustomDialog.alert("Administrators haven not configured the external repository.", "Unable to upload file");
        }

        // no file was selected
        if (!items || items.length === 0) return;

        // parse selected file
        items.forEach( async (item) => {
            // if item already has a URL, item is already uploaded.
            if (item.fileAbsoluteUrl) return this.updateFields(item);

            // upload this file
            const uploadItem = await this.uploadItem(item);
            if (uploadItem) {
                item.fileAbsoluteUrl = (new URL((uploadItem as Record<string, any>).ServerRelativeUrl, location.origin)).href;
                item.fileName = (uploadItem as Record<string, any>).Name;

                this.updateFields(item);
                
            } else {
                // something went wrong with the upload
                await CustomDialog.alert("Unable to upload file. View console for more details.", "ERROR!");
            }

        });
    }
    //#endregion

}
//#endregion

//#region RICH TEXT EDITOR - RTE (uses TinyMCE React)
export interface IRichTextControlProps extends IRichTextProps {
    required?: boolean;
    id?: string;
    label?: string;
    description?: string;
    errorMessage?: string;
    disabled: boolean;
}
export class RichTextControl extends React.Component<IRichTextControlProps> {
    public render(): React.ReactElement<IRichTextControlProps> {
        const props: IRichTextControlProps = assign({}, this.props, {
            isEditMode: !this.props.disabled,

            // remove props that are not RichText
            label: undefined,
            required: undefined,
            description: undefined,
            errorMessage: undefined
        });

        let labelClassNames = [],
            richTextWrapperClassNames = [styles.richTextWrapper];

        // show red border if error occured
        if (this.props.errorMessage) richTextWrapperClassNames.push( styles.richTextWrapperError );

        // label classnames
        labelClassNames.push("ms-label");
        if (this.props.required) labelClassNames.push(styles.required);

        // NOTE: RichText control needs to be expressed verbosely in order for the 'value' prop to be accepted.
        return (
            <FieldWrapper>
                <LabelElement htmlFor={ this.props.id } required={ this.props.required }>{ this.props.label }</LabelElement>
                <div className={ richTextWrapperClassNames.join(" ") }>
                    { React.createElement(RichText, props) }
                </div>
                <DescriptionElement text={this.props.description}/>
                <ErrorMessageElement text={this.props.errorMessage}/>
            </FieldWrapper>
        );
    }
}
//#endregion

//#region DATE TIME CONTROL
export interface IDateTimeControlProps extends IDateTimePickerProps {
    id: string;
    label?: string;
    required?: boolean;
    description?: string;
    errorMessage?: string;
}
export class DateTimeControl extends React.Component<IDateTimeControlProps> {
    public render(): React.ReactElement<IDateTimeControlProps> {
        const props: IDateTimePickerProps = assign({}, this.props, {
            key: this.props.id,

            // remove props that are not IDateTimePickerProps
            id: undefined,
            label: undefined,
            required: undefined,
            description: undefined,
            errorMessage: undefined
        });

        return (
            <FieldWrapper>
                <LabelElement htmlFor={ this.props.id } required={ this.props.required }>{ this.props.label }</LabelElement>
                { React.createElement(DateTimePicker, props) }
                <DescriptionElement text={this.props.description}/>
                <ErrorMessageElement text={this.props.errorMessage}/>
            </FieldWrapper>
        );
    }
}
//#endregion

//#region COMBOBOXLISTITEMPICKER, LISTITEMPICKER: LOOKUP
// dropdown control
export interface IComboBoxListItemPickerControlProps extends IComboBoxListItemPickerProps {
    description?: string;
    errorMessage?: string;
    required?: boolean;
    label?: string;
    //selectedItems?: any[];
}
export class ComboBoxListItemPickerControl extends React.Component<IComboBoxListItemPickerControlProps> {
    public render(): React.ReactElement<IComboBoxListItemPickerControlProps> {
        const props = assign({}, this.props, {
            // remove props that are not ComboBoxListItemPicker
            description: undefined,
            errorMessage: undefined,
            required: undefined,
            label: undefined
        });

        return (
            <FieldWrapper>
                <LabelElement required={ this.props.required }>{ this.props.label }</LabelElement>
                {/* { React.createElement(ComboBoxListItemPicker, props) } */}
                <ComboBoxListItemPicker
                    disabled={ this.props.disabled }
                    spHttpClient={ this.props.spHttpClient }
                    webUrl={ this.props.webUrl }
                    listId={ this.props.listId }
                    columnInternalName={ this.props.columnInternalName }
                    multiSelect={ this.props.multiSelect }
                    defaultSelectedItems={ this.props.defaultSelectedItems }
                    onSelectedItem={ this.props.onSelectedItem }
                />
                <DescriptionElement text={this.props.description}/>
                <ErrorMessageElement text={this.props.errorMessage}/>
            </FieldWrapper>
        );
    }
}

// type and select box
export interface IListItemPickerControlProps extends IListItemPickerProps {
    required?: boolean;
    description?: string;
    errorMessage?: string;
}
export class ListItemPickerControl extends React.Component<IListItemPickerControlProps> {
    public render(): React.ReactElement<IListItemPickerControlProps> {
        const props = assign({}, this.props, {
            // remove props that are not IListItemPickerProps
            label: undefined,
            required: undefined,
            description: undefined,
            errorText: undefined
        });
        return (
            <FieldWrapper>
                <LabelElement required={ this.props.required }>{ this.props.label }</LabelElement>
                { React.createElement(ListItemPicker, props) }
                <DescriptionElement text={this.props.description}/>
                <ErrorMessageElement text={this.props.errorMessage}/>
            </FieldWrapper>
        );
    }
}
//#endregion

//#region DROPDOWN
export interface IDropdownControlProps extends IDropdownProps {
    description?: string;
}
export class DropdownControl extends React.Component<IDropdownControlProps> {
    public render(): React.ReactElement<IDropdownControlProps> {
        const description = this.props.description || null,
            props = assign({}, this.props, { description: undefined });

        return (
            <FieldWrapper>
                { React.createElement(Dropdown, props) }
                { React.createElement(DescriptionElement, { text: description }) }
            </FieldWrapper>
        );
    }
}
//#endregion

//#region PEOPLE PICKER CONTROL
export interface IPeoplePickerControlProps extends IPeoplePickerProps {
    label?: string;
    description?: string;
}
export class PeoplePickerControl extends React.Component<IPeoplePickerControlProps> {
    public render(): React.ReactElement<IPeoplePickerControlProps> {
        const props: IPeoplePickerProps = assign({}, this.props, {
            // remove props that are not IPeoplePickerProps
            label: undefined,
            description: undefined
        });

        return (
            <FieldWrapper>
                <LabelElement required={ this.props.required }>{ this.props.label }</LabelElement>
                { React.createElement(PeoplePicker, props)}
                <DescriptionElement text={this.props.description}/>
            </FieldWrapper>
        );
    }
}
//#endregion

//#region COMPLEX DATA CONTROL
export interface ICustomControlComplexDataProps extends AboutUsDisplay.IAboutUsComplexDataDisplayProps {
    displayControl: React.ComponentClass<any>;
    label: string;

    disabled?: boolean;
    required?: boolean;
    description?: string;
    errorMessage?: string;

    onAdd?: ()=>any;
}

export class CustomControlComplexData extends React.Component<ICustomControlComplexDataProps> {
    public render(): React.ReactElement<ICustomControlComplexDataProps> {
        const addButtonProps: IButtonProps = {
                iconProps: { iconName: "Add" },
                text: `Add ${ this.props.label }`,
                disabled: this.props.disabled,
                onClick: this.props.onAdd
            },
            displayProps: AboutUsDisplay.IAboutUsComplexDataDisplayProps = {
                properties: this.props.properties,
                values: this.props.values,
                showEditControls: (this.props.disabled === true) ? false : true,
                onOrderChange: this.props.onOrderChange,
                onEdit: this.props.onEdit,
                onDelete: this.props.onDelete,
                extraButtons: this.props.extraButtons
            };

        return (
            <FieldWrapper>
                <LabelElement required={ this.props.required }>{ this.props.label }</LabelElement>
                <div>{ (this.props.disabled !== true) ? React.createElement(ActionButton, addButtonProps) : null }</div>
                <DescriptionElement text={ this.props.description }/>
                <ErrorMessageElement text={ this.props.errorMessage }/>
                <div className={ styles.complexDataDisplayContainer }>
                    { React.createElement(this.props.displayControl, displayProps) }
                </div>
            </FieldWrapper>
        );
    }
}
//#endregion

//#region TAGS & KEYWORDS CONTROL
export interface ICustomControlKeywordsProps extends AboutUsDisplay.IAboutUsKeywordsDisplayProps {
    label: string;

    disabled?: boolean;
    required?: boolean;
    description?: string;
    errorMessage?: string;

    onAdd?: (value: string)=>ICustomControlKeywordsState;
}

export interface ICustomControlKeywordsState {
    value: string;
}

export class CustomControlKeywords extends React.Component<ICustomControlKeywordsProps, ICustomControlKeywordsState> {
    constructor(props) {
        super(props);

        this.state = {
            value: ""
        };
    }
    public render(): React.ReactElement<ICustomControlKeywordsProps> {
        const addButtonProps: IButtonProps = {
                iconProps: { iconName: "Add" },
                disabled: this.props.disabled,
                onClick: () => { this.setState(this.props.onAdd(this.state.value)); }
            },
            displayProps: AboutUsDisplay.IAboutUsKeywordsDisplayProps = {
                values: this.props.values,
                showEditControls: (this.props.disabled === true) ? false : true,
                onOrderChange: this.props.onOrderChange,
                onDelete: this.props.onDelete
            };

        return (
            <FieldWrapper>
                <LabelElement required={ this.props.required }>{ this.props.label }</LabelElement>
                <div className={ styles.textboxWithButtonWrapper }>
                    <input
                        className={ styles.textbox }
                        disabled={this.props.disabled}
                        value={ this.state.value }
                        onChange={ evt => this.setState({"value": evt.target.value}) }
                        onKeyPress={ evt => {
                            if (evt.keyCode === 13 || evt.which === 13) this.setState(this.props.onAdd(this.state.value)); 
                        } } />
                    { (!this.props.disabled) ? 
                        <TooltipHost content="Add">
                            { React.createElement(IconButton, addButtonProps) }
                        </TooltipHost>
                        : null
                    }
                </div>
                <DescriptionElement text={ this.props.description }/>
                <ErrorMessageElement text={ this.props.errorMessage }/>
                <div className={ styles.complexDataDisplayContainer }>
                    { React.createElement(AboutUsDisplay.KeywordsDisplay, displayProps) }
                </div>
            </FieldWrapper>
        );
    }
}
//#endregion
