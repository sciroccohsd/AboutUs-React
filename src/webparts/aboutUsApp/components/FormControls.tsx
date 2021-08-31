import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { assign } from 'lodash';
// https://docs.microsoft.com/en-us/javascript/api/office-ui-fabric-react?view=office-ui-fabric-react-latest
import { Dropdown, 
    IDropdownProps,
    ILabelProps,
    ISpinnerProps,
    ITextFieldProps, 
    Label, 
    Spinner, 
    TextField } from 'office-ui-fabric-react';

// https://pnp.github.io/sp-dev-fx-controls-react/index.html
//> npm install @pnp/spfx-controls-react --save
import * as ReactControls from '@pnp/spfx-controls-react';
import { ComboBoxListItemPicker, 
    DateTimePicker, 
    IComboBoxListItemPickerProps, 
    IDateTimePickerProps, 
    IListItemPickerProps, 
    IPlaceholderProps, 
    ListItemPicker } from '@pnp/spfx-controls-react';
import { IRichTextProps, RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { PeoplePicker, PrincipalType, IPeoplePickerProps } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { Toolbar } from '@pnp/spfx-controls-react/lib/Toolbar';
import { IFieldUrlValue } from './DataFactory';

export default ReactControls;
export * from '@pnp/spfx-controls-react/lib/RichText';
export * from '@pnp/spfx-controls-react/lib/PeoplePicker';
export * from '@pnp/spfx-controls-react/lib/Toolbar';


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
                label: "loading...",
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
        const props = assign({}, this.props, { className: "FormControlsLabel "  + this.props.className || ""});
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
                    <span className="FormControlsDescription">
                        <span className={ `ms-TextField-description ${ styles.description }`}>
                            { this.props.text }
                        </span>
                    </span>
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
}
export class UrlControl extends React.Component<IUrlControlProps> {
    public render(): React.ReactElement<IUrlControlProps> {
        const showTextField = typeof this.props.showTextField === "boolean" ? this.props.showTextField : true ,
            urlValue = this.props.defaultValue ? this.props.defaultValue.Url : null,
            textValue = this.props.defaultValue ? this.props.defaultValue.Description : urlValue,
            onChange = this.props.onChange ? evt => {
                const isUrl = evt.target.id === this.props.id,
                    value = evt.target.value;

                this.props.onChange.call(this.props.onChange, value, (isUrl)?"url":"text", this.props.id);
            } : null,
            urlProps: ITextFieldProps = {
                type: "url",
                id: this.props.id,
                className: `FormControlsUrlValue ${this.props.className || ""}`,
                label:  this.props.label,
                defaultValue: urlValue,
                required: this.props.required,
                disabled: this.props.disabled,
                placeholder: "https://...",

                onChange: onChange
            },
            textProps: ITextFieldProps = {};

        if (showTextField) {
            // show both URL and Text fields, populate properties for text field
            textProps.disabled = this.props.disabled;
            textProps.id = this.props.id + "_text";
            textProps.className = `FormControlsTextValue ${this.props.className || ""}`;
            textProps.label = `Text for: ${ this.props.label || "Url"}`;
            textProps.defaultValue = textValue;
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
           </FieldWrapper>
        );
    }
}
//#endregion

//#region RICH TEXT EDITOR - RTE (uses TinyMCE React)
export interface IRichTextControlProps extends IRichTextProps {
    required?: boolean;
    id?: string;
    label?: string;
    description?: string;
    errorMessage?: string;
    disabled?: boolean;
}
export class RichTextControl extends React.Component<IRichTextControlProps> {
    public render(): React.ReactElement<IRichTextControlProps> {
        let labelClassNames = [],
            richTextWrapperClassNames = [styles.richTextWrapper];
        const props: IRichTextProps = assign({}, this.props, {
                className: `${(this.props.className || "")} ${ styles.richtext }`,

                // remove props that are not IRichTextProps
                required: undefined,
                id: undefined,
                label: undefined,
                description: undefined,
                errorMessage: undefined,
                disabled: undefined
            });

        // show red border if error occured
        if (this.props.errorMessage) richTextWrapperClassNames.push( styles.richTextWrapperError );

        // props.disabled & props.isEditMode is the same property. isEditMode value takes priority.
        props.isEditMode = this.props.isEditMode === false 
            ? false 
            : this.props.disabled ? false : true;

        // label classnames
        labelClassNames.push("ms-label");
        labelClassNames.push(styles.label);
        if (this.props.required) labelClassNames.push(styles.required);

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
                { React.createElement(ComboBoxListItemPicker, props) }
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