// SP-Dialog extension for modal alerts, prompts and custom modal forms.
import { BaseDialog, IDialogShowOptions } from '@microsoft/sp-dialog';
import styles from './CustomDialog.module.scss';

//#region INTERFACES & ENUMS
export interface ICustomDialogBody {
    "wrapper": HTMLDivElement;
    "container": HTMLDivElement;
    "actions": HTMLDivElement;
}
export interface ICustomDialogField {
    "wrapper": HTMLDivElement;
    "label": HTMLLabelElement;
    "description": HTMLDivElement;
    "error": HTMLDivElement;
    "fields": HTMLInputElement[] | HTMLTextAreaElement[] | HTMLSelectElement[] | any[];
}
//#endregion

/** Use the premade custom methods (returns Promise) or create your own custom dialog. 
 * - alert(msg, title, showClose); 
 * - prompt(msg, title, defaultValue, label, showClose); 
 * - confirm(msg, title, {yes: "Yes", no: "No"}, showClose); 
 * 
 * @remarks
 * Extends 'BaseDialog' from \@microsoft/sp-dialog.
 * 
 * @example 
 * // alert()
 * await CustomDialog.alert("I'm an Alert", "Custom Title");
 * 
 * // prompt()
 * const str = await CustomDialog.prompt("I'm a prompt!", "Custom Title", "Default value", "Label:", true);
 * 
 * // confirm()
 * const bln = await CustomDialog.confirm("I'm a confirm message!", "Custom Title", null, true);
 */
export default class CustomDialog extends BaseDialog {
    //#region PROPERTIES
    private _root: HTMLDivElement = this.CreateElement("div", { "class": styles.customDialog });
    private _header: HTMLDivElement;
    private _body: ICustomDialogBody;
    private _form: HTMLDivElement = null;
    private _fieldsList: {} = {};
    private _actionsList: {} = {};
    private _onSubmit: (evt)=>any = (evt) => { this.close(); return false; };
    //#endregion

    //#region GETTERS & SETTERS
    // root
    public get root(): HTMLDivElement {
        return this._root;
    }  

    // header
    public get header(): HTMLDivElement {
        return this._header;
    }
    
    // body
    public get body(): HTMLDivElement {
        return this._body.container;
    }
    public get actions(): HTMLDivElement {
        return this._body.actions;
    }
    
    // fieldGroup
    public get form(): HTMLDivElement {
        return this._form;
    }
    //#endregion

    //#region CONSTRUCTOR
    /** Create a new custom sp-dialog.
     * @param title Title for dialog box. Keep it short.
     * @param showClose Show/hide the 'X' close button.
     * @example
     * // custom dialog
     * const dialog = new CustomDialog("Custom Title", true);
     * await dialog.AddMessage("I'm a custom dialog!").AddCancelAction("Yep!").show();
     */
    constructor(title: string, showClose: boolean = false, formAction?: (evt)=>{}) {
        super();
        this._header = this.CreateDialogHeader(title, showClose);
        this._body = this.CreateDialogBody(formAction);
    }
    //#endregion

    //#region PUBLIC STATIC METHODS: ALERT, PROMPT, CUSTOM
    /** Show a custom SP alert.
     * @param msg Text to display.
     * @param title Dialog title. Default: 'Message'
     * @param showClose Show/hide the 'X' close button. Default: false
     * @returns Promise => true when closed
     * @example
     * Example:
     * const result = await CustomDialog.alert("Hello!");
     */
    public static async alert(msg: string, title: string = "Message", showClose: boolean = false): Promise<boolean> {
        const eDialog = new CustomDialog(title, showClose);

        await eDialog.AddMessage(msg)
            .AddCancelAction("OK")
            .show();

        return true;
    }

    /** Show a custom SP modal message. Use .close() to close the dialog.
     * @param msg Text to display.
     * @param title Dialog title. Default: 'Message'
     * @returns CustomDialog object.
     * @example
     * Example:
     * const modalMsg = CustomDialog.modalMsg("Hello!");
     * // do stuff...
     * modalMsg.close();
     */
    public static modalMsg(msg: string, title: string = "Message"): CustomDialog {
        const eDialog = new CustomDialog(title);

        eDialog.AddMessage(msg).show();

        return eDialog;
    }

    /** Show custom SP prompt (text response dialog).
     * @param msg Text to display.
     * @param title Dialog title. Default: 'Prompt'
     * @param opt Prompt options:
     * - value: string = Default value for textbox.
     * - label: string = Label for textbox.
     * - description: string = HTML to display as the fields description
     * - error: string = Error text to display below textbox.
     * @param showClose Show/hide the 'X' close button. Default: false
     */
    public static async prompt(msg: string, title: string = "Prompt", opt: {} = {}, showClose: boolean = false): Promise<string> {
        // ensure options
        if (typeof opt !== "object" || opt === null) opt = {};
        if (!("value" in opt)) opt["value"] = "";
        if (!("label" in opt)) opt["label"] = "";
        if (!("description" in opt)) opt["description"] = "";
        if (!("error" in opt)) opt["error"] = "";
        
        let value: string = null;
        const dialog = new CustomDialog(title, showClose),
            id = `dialog-textfield-${ dialog.RandomNumber() }`;

        await dialog.AddMessage(msg)
            .AddTextField(opt["label"], id, { "value": opt["value"], "error": opt["error"], "description": opt["description"] })
            .AddCancelAction("Close", { "tabIndex": 0 })
            .AddAction("OK", {
                "tabIndex": 1,
                "onclick": evt => {
                    value = dialog.GetFieldElements(id).fields[0].value;
                    dialog.close();
                }
            }, true)
            .show();

        return value;
    }

    /** Show custom SP confirm (Yes/No dialog).
     * @param msg Text to display.
     * @param title Dialog title. Default: 'Confirm?'
     * @param yesText Text for 'Yes' button. Default: 'Yes'
     * @param noText Text for 'No' button. Default: 'No'
     * @param showClose Show/hide the 'X' close button. Default: false
     * @returns Promise => true, false or null.
     */
    public static async confirm(msg: string, title: string = "Confirm?", text: {"yes": string, "no": string} = {"yes": "Yes", "no": "No"}, showClose: boolean = false): Promise<boolean> {
        // ensure yes/no text
        if (typeof text !== "object" || text === null) text = {"yes": null, "no": null};
        if (!("yes" in text) || typeof text["yes"] !== "string") text["yes"] = "Yes";
        if (!("no" in text) || typeof text["no"] !== "string") text["no"] = "No";
        
        let value: boolean = null;
        const dialog = new CustomDialog(title, showClose);

        await dialog.AddMessage(msg)
            .AddAction(text.no, {
                "tabIndex": 0,
                "onclick": async evt => {
                    value = false;
                    await dialog.close();
                }
            })
            .AddAction(text.yes, {
                "tabIndex": 1,
                "onclick": async evt => {
                    value = true;
                    await dialog.close();
                }
            }, true)
            .show();

        return value;
    }
    //#endregion

    //#region ELEMENTS: MAIN SECTIONS (HEADER, BODY, SUBTEXT, FORM)
    /** Generates the dialog header section with a DIV wrapper.
     * Ready to append to dialog container.
     * @param title Title text. Keep it short.
     * @param showClose Show/hide the 'X' close button. Default: false.
     * @returns DIV element
     */
    private CreateDialogHeader(title: string = "", showClose: boolean = false): HTMLDivElement {
        const eHeader: HTMLDivElement = this.CreateElement("div", { "class": `ms-Dialog-header ${ styles.header }` }),
            eHeaderTitle: HTMLDivElement = this.CreateElement("div", { "class": `ms-Dialog-title ${ styles.title }`, "role": "heading", "text": title }),
            eTopButton: HTMLDivElement = this.CreateElement("div", { "class": `${ styles.topButtonWrapper }` });

        // show "X" close button: entity = "&#9932;"
        if (showClose) {
            const cancelButton = this.CreateElement("button", {
                    "type": "button",
                    "class": styles.topButton,
                    "title": "Close",
                    "onclick": this.close
                }),
                cancelIcon = this.CreateIcon("&#9932;");
                //cancelIcon = this.CreateIcon("Cancel");

            cancelButton.append(cancelIcon);
            eTopButton.append(cancelButton);
        }

        eHeader.append(eHeaderTitle, eTopButton);

        return eHeader;
    }
    
    /** Generates the dialog body sections with a DIV wrapper.
     * Ready to append to the dialog container.
     * @returns DIV wrapper, DIV body & DIV actions (actions)
     */
     private CreateDialogBody(formAction?: (evt)=>{}): ICustomDialogBody {
        const eWrapper: HTMLDivElement = this.CreateElement("div", { "class": `ms-Dialog-inner ${ styles.inner }` }),
            eForm: HTMLFormElement = this.CreateElement("form", { "action": "javascript:void(0);" }),
            eContainer: HTMLDivElement = this.CreateElement("div", { "class": `ms-Dialog-content ${ styles.innerContent }` }),
            eActions: HTMLDivElement = this.CreateElement("div", { "class": `ms-Dialog-actions ${ styles.actions }` }),
            eActionsRight: HTMLDivElement = this.CreateElement("div", { "class": `ms-Dialog-actionsRight ${ styles.actionsRight }` });

        // put elements together
        eWrapper.append(eForm);
        eForm.append(eContainer, eActions);
        eActions.append(eActionsRight);

        // add form action
        eForm.addEventListener("submit", evt => { this._onSubmit.call(this, evt); }, false);

        return {"wrapper": eWrapper, "container": eContainer, "actions": eActionsRight};
    }

    /** Generates a text section in a DIV element.
     * Ready to append to body container.
     * @param msg Text to display.
     * @returns DIV element
     */
     private CreateDialogSubText(msg: string = ""): HTMLDivElement {
        const eMessage: HTMLDivElement = this.CreateElement("div", { "class": styles.subText, "text": msg });
        return eMessage;
    }

    /** Create a Fluent UI Icon (i) element or HTML Entity (span) element.
     * @param iconName Name of Fluent UI Icon or HTML Entity. Case-sensitive.
     * - https://developer.microsoft.com/en-us/fluentui#/styles/web/icons
     * - https://www.w3schools.com/charsets/ref_html_entities_a.asp
     * @returns I Element
     */
    private CreateIcon(iconName): HTMLElement {
        // Download Fluent UI Icons from: https://uifabricicons.azurewebsites.net/
        // - Place fabric-icons-inline.scss (rename to fabric-icons-inline.module.scss) in the webpart\[PROJECT] folder
        // - Then @import './fabric-icons-inline.module.scss'; in the projects .scss file.

        /// ** Uncomment to use Fluent UI Icons.
        //const icon = this.CreateElement("i", { "class": `${ styles['ms-Icon'] } ${ styles[`ms-Icon--${ iconName }`] }`, "aria-hidden": "true" });
        
        /// ** or Uncomment to use HTML Entity names
        const icon = this.CreateElement("span", { "html": iconName });

        return icon;
    }
    
    /** Generates the fields container with a DIV wrapper.
     * Ready to append to the body container.
     * Use this container to add fields.
     * @returns DIV root & DIV container
     */
     private CreateDialogForm(): HTMLDivElement {
        const eForm: HTMLDivElement = this.CreateElement("div", { "class": styles.form });
        return eForm;
    }
    //#endregion

    //#region ELEMENTS: FORM FIELDS (INPUT:TEXT, SELECT, TEXTAREA...)
    /** Create a INPUT:Text field with a LABEL wrapper.
     * Ready to append to the field group container.
     * @param label Text for field label.
     * @param id Unique element 'ID' for the field.
     * @param opt Other options
     * - name: string = Form field name. Uses the fieldId if 'name' is not provided.
     * - value: string = Default value
     * - hint: string = Hoverover text
     * - description: string = HTML to display as the fields description
     * - error: string = Error text
     * - class: string = CSS class(es) to append to the wrapper
     * @returns ICustomDialogField
     */
    private CreateTextField(label: string = "", id: string, opt: {} = {}): Partial<ICustomDialogField> {
        // ensure options
        if (typeof opt !== "object" || opt === null) opt = {};
        if (!("name" in opt)) opt["name"] = id;
        if (!("value" in opt)) opt["value"] = "";
        if (!("hint" in opt)) opt["hint"] = "";
        if (!("error" in opt)) opt["error"] = "";
        if (!("description" in opt)) opt["description"] = "";
        if (!("class" in opt)) opt["class"] = "";

        const eWrapper = this.CreateElement("div", { "class": `${ styles.fieldWrapper } ${ opt["class"] }` }),
            eLabel = this.CreateElement("label", { "for": id, "class": styles.fieldLabel, "text": label }),
            eDescription = this.CreateElement("div", { "class": styles.fieldDescription, "html": opt["description"] }),
            eError = this.CreateElement("div", { "class": styles.fieldError, "text":  opt["error"] }),
            eField: HTMLInputElement = this.CreateElement("input", { 
                "type": "text",
                "id": id,
                "name": opt["name"],
                "value": opt["value"],
                "class": `ms-TextField-field ${ styles.field }`,
                "title": opt["hint"]
            });

        eWrapper.append(eLabel, eField, eDescription, eError);

        return {
            "wrapper": eWrapper,
            "label": eLabel,
            "description": eDescription,
            "error": eError,
            "fields": [eField]
        };
    }


    // TODO: Create the rest of the form fields: checkbox, radio, textarea, select
    //#endregion

    //#region ELEMENTS: DIALOG ACTION BUTTONS
    /** Generates a dialog action button in a SPAN wrapper.
     * Ready to append to the Actions container.
     * @param label Text displayed on button.
     * @param opt HTML attributes for BUTTON element. Common attributes to add: 'tabOrder' & 'onclick'.
     * @returns SPAN element
     */
    private CreateActionButton(label: string, opt: {[key:string]: any} = {}, isPrimaryButton: boolean = false): HTMLSpanElement {
        // generate css class for the Button
        let css: string[] = ["ms-Button"];

        // add default customDialog class
        css.push(styles.buttonRoot);
        if (isPrimaryButton) css.push(styles.buttonPrimary);

        // append user-defined classes
        if ("class" in opt) css.push(opt["class"]);

        opt["class"] = css.join(" ");

        // add element type
        if (!("type" in opt)) opt["type"] = "button";

        // remove "text" property from attr.
        if ("text" in opt) delete opt["text"];

        const eWrapper: HTMLSpanElement = this.CreateElement("span", { "class": `ms-Dialog-action ${ styles.action }`}),
            eButton: HTMLButtonElement = this.CreateElement("button", opt),
            eContainer: HTMLSpanElement = this.CreateElement("span", { "class": `ms-Button-flexContainer ${ styles.buttonContainer }`}),
            eLabelContainer: HTMLSpanElement = this.CreateElement("span", { "class": `ms-Button-textContainer ${ styles.buttonLabelContainer }` }),
            eLabel: HTMLSpanElement = this.CreateElement("span", { "class": `ms-Button-label ${ styles.buttonLabel }`, "text": label });

        eWrapper.append(eButton);
        eButton.append(eContainer);
        eContainer.append(eLabelContainer);
        eLabelContainer.append(eLabel);

        return eWrapper;
    }

    /** Generates a dialog submit button in a SPAN wrapper.
     * Ready to append to the Actions container.
     * @param onSubmit Form action event listener
     * @param label Text displayed on button.
     * @param opt HTML attributes for BUTTON element. Common attributes to add: 'tabOrder' & 'onclick'.
     * @returns SPAN element
     */
    private CreateSubmitButton(onSubmit: (evt)=>any = null, label: string, opt: {[key:string]: any} = {}, isPrimaryButton: boolean = true): HTMLSpanElement {
        if (onSubmit && typeof onSubmit === "function") this._onSubmit = onSubmit;
        
        // generate css class for the Button
        let css: string[] = ["ms-Button"];

        // add default customDialog class
        css.push(styles.buttonRoot);
        if (isPrimaryButton) css.push(styles.buttonPrimary);

        // append user-defined classes
        if ("class" in opt) css.push(opt["class"]);

        opt["class"] = css.join(" ");

        // add element type
        if (!("type" in opt)) opt["type"] = "submit";
        if (!("value" in opt)) opt["value"] = label || "submit";

        // remove "text" property from attr.
        if ("text" in opt) delete opt["text"];

        const eWrapper: HTMLSpanElement = this.CreateElement("span", { "class": `ms-Dialog-action ${ styles.action }`}),
            eButton: HTMLInputElement = this.CreateElement("input", opt);

        eWrapper.append(eButton);

        return eWrapper;
    }

    /** Generates a standard 'OK' dialog action button in a wrapper that closes the dialog.
     * Ready to append to the Actions container.
     * @param label Text displayed on button. Default: 'OK'.
     * @param opt HTML attributes for BUTTON element. 'onclick' is predefined.
     * @returns SPAN element
     */
    private CreateCloseActionButton(label: string = "OK", opt: {} = {}): HTMLSpanElement {
        opt["onclick"] = this.close;
        return this.CreateActionButton(label, opt);
    }
    //#endregion

    //#region CUSTOMDIALOG METHODS
    /** Add a text (DIV) section to this dialog.
     * @param msg Text to display.
     * @returns self
     */
    public AddMessage(msg: string): CustomDialog {
        const eMessage = this.CreateDialogSubText(msg);
        this._body.container.append(eMessage);
        return this;
    }

    /** Add an action button to this dialog.
     * @param label Text displayed on button.
     * @param opt HTML attributes for BUTTON element. Common attributes to add: 'tabOrder' & 'onclick'.
     * @returns self
     */
    public AddAction(label: string, opt: {} = {}, isPrimaryButton: boolean = false): CustomDialog {
        const eButton = this.CreateActionButton(label, opt, isPrimaryButton);
        this._body.actions.append(eButton);
        return this;
    }

    /** Add a input:submit button to this dialog.
     * @param onSubmit Form action event listener
     * @param label Text displayed on button.
     * @param opt HTML attributes for BUTTON element. Common attributes to add: 'tabOrder' & 'onclick'.
     * @returns self
     */
    public AddSubmit(onSubmit: (evt)=>any = null, label: string = "Submit", opt: {} = {}, isPrimaryButton: boolean = true): CustomDialog {
        const eButton = this.CreateSubmitButton(onSubmit, label, opt, isPrimaryButton);
        this._body.actions.append(eButton);
        return this;
    }

    /** Add a 'Close' action button to this dialog.
     * @param label Text displayed on button. Default: 'Cancel'.
     * @param opt HTML attributes for BUTTON element. 'onclick' is predefined.
     * @returns self
     */
    public AddCancelAction(label: string = "Cancel", opt: {} = {}): CustomDialog {
        const eButton = this.CreateCloseActionButton(label, opt);
        this._body.actions.append(eButton);
        return this;
    }

    /** Add a text INPUT field to the Field Group. Adds the Field Group if not already created. 
     * @param label Text for field label.
     * @param id Unique element 'ID' for the field.
     * @param opt Other options
     * - name: string = Form field name. Uses the fieldId if 'name' is not provided.
     * - value: string = Default value
     * - hint: string = Hoverover text
     * - description: string = HTML to display as the fields description
     * - class: string = CSS class(es) to append to the wrapper
     * @returns self
     */
    public AddTextField(label: string = "", id: string, opt: {} = {}): CustomDialog {
        // create field group if it doesn't exist
        if (this._form === null) {
            this._form = this.CreateDialogForm();
            this._body.container.append(this._form);
        }

        const eField = this.CreateTextField(label, id, opt);

        // warn if this is a duplicate ID. user may or may not have wanted to overwrite it.
        if (id in this._fieldsList) {
            console.warn(`Overwritten: Field already exists in custom dialog template! Field ID: '${ id }'`);
        }

        this._fieldsList[id] = eField;
        this.form.append(eField.wrapper);

        return this;
    }

    /** Get the form field elements that was added for a specific 
     * @param id_or_index ID or index number for field.
     * @returns Wrapper (DIV), label (Label) & field (input|select|textarea...)
     */
    public GetFieldElements(id_or_index?: string | number): ICustomDialogField {
        const keys = Object.keys(this._fieldsList);

        if (typeof id_or_index === "string") {
            return (id_or_index in this._fieldsList) ? this._fieldsList[id_or_index] : null;
        } else if (typeof id_or_index === "number") {
            return (id_or_index > -1 && id_or_index < keys.length) ? this._fieldsList[keys[id_or_index]] : null;
        } else {
            return (keys.length > 0) ? this._fieldsList[keys[keys.length - 1]] : null;
        }

    }
    //#endregion

    //#region HELPERS
    /** Create a HTML DOM element.
     * @param name HTML tag name. Example: 'div', 'span', 'button', 'p', ...
     * @param attr Standard HTML element attributes. Shorthand: 'text' = 'innerText'; 'html' = 'innerHTML'; 'class' = 'className'.
     * @returns HTML DOM element
     */
    public CreateElement(name: string, attr: {} = {}): HTMLElement | any {
        const elem = document.createElement(name);

        // add attributes
        for (const key in attr) {
            let value = attr[key];
            
            switch (key) {
                case "text":
                    elem.innerText = value;
                    break;
                case "html":
                    elem.innerHTML = value;
                    break;
                case "class":
                    elem.className = value;
                    break;
            
                default:
                    elem[key] = value;
                    break;
            }
                
        }

        return elem;
    }
    
    /** Generate a random positive integer.
     * @param min Minimum random number.
     * @param max Maximum random number.
     * @returns Random integer
     */
    public RandomNumber(min: number = 10000, max: number = 99999): number {
        min = Math.min(Math.abs(min), Math.abs(max));
        max = Math.max(Math.abs(min), Math.abs(max));
        return Math.floor(Math.random() * (max - min)) + min;
    }
    //#endregion

    //#region RENDER
    public render(): void {
        this._root.append(this._header, this._body.wrapper);
        this.domElement.append(this._root);
    }

    protected onAfterClose(): void {
        super.onAfterClose();
        this._root.innerHTML = "";
    }
    //#endregion
}