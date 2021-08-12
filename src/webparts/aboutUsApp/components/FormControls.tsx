import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { assign } from 'lodash';
// https://docs.microsoft.com/en-us/javascript/api/getstarted/getstartedpage?view=office-ui-fabric-react-latest
import { Dropdown, IComboBoxProps, IDropdown, IDropdownProps, IInputProps, ILabelProps, IStyle, ITextFieldProps, TextField } from 'office-ui-fabric-react';

// https://docs.microsoft.com/en-us/javascript/api/getstarted/getstartedpage?view=office-ui-fabric-react-latest
//> npm install @fluentui/react
//import { ComboBox } from '@fluentui/react';

// https://pnp.github.io/sp-dev-fx-controls-react/index.html
//> npm install @pnp/spfx-controls-react --save
import * as ReactControls from '@pnp/spfx-controls-react';
import { IPlaceholderProps } from '@pnp/spfx-controls-react';
import { IRichTextProps, RichText } from '@pnp/spfx-controls-react/lib/RichText';

// https://github.com/sstur/react-rte
//> npm install --save react-rte
//import RichTextEditor from 'react-rte';

// https://www.tiny.cloud/docs/integrations/react/
//> npm install --save @tinymce/tinymce-react
// import { Editor } from '@tinymce/tinymce-react';

// https://draftjs.org/docs/getting-started
//> npm install draft-js babel-polyfill --save
// import {Editor, EditorState} from 'draft-js';
// import 'draft-js/dist/Draft.css';

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
                <label htmlFor={ this.props.id } className={ labelClassNames.join(" ") }>{ this.props.label }</label>
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

//#region DROPDOWN
export interface IFormControlDropdownProps extends IDropdownProps {
    description?: string;
}
export class DropdownControl extends React.Component<IFormControlDropdownProps> {
    public render(): React.ReactElement<IFormControlDropdownProps> {
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