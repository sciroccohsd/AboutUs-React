// About-Us data display elements.
import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { assign, divide, find, trim } from 'lodash';

// https://docs.microsoft.com/en-us/javascript/api/office-ui-fabric-react?view=office-ui-fabric-react-latest
import { CommandBar, Dropdown,
    FontIcon,
    IButtonProps,
    ICommandBarItemProps,
    IconButton,
    IDropdownProps,
    ILabelProps,
    ISpinnerProps,
    ITextFieldProps, 
    ITooltipProps, 
    Label, 
    Spinner, 
    Stack, 
    TextField, 
    TooltipHost} from 'office-ui-fabric-react';
import { IAboutUsValueState } from './AboutUsForm';

//> npm install react-easy-sort
import SortableList, { SortableItem } from 'react-easy-sort';
import AboutUsAppWebPart, { IAboutUsAppWebPartProps } from '../AboutUsAppWebPart';
import { Wrapper } from './AboutUsApp';
import DataFactory, { IDataStructureItem } from './DataFactory';
import { WebPartContext } from '@microsoft/sp-webpart-base';


//#region COMPLEX DATA DISPLAYS
    //#region COMPLEX DATA INTERFACE
    export type TAboutUsComplexData = Record<string, any>;

    interface IAboutUsComplexDataCommandBarProps {
        itemIndex: number;
        onEdit?: (index: number)=>any;
        onDelete?: (index: number)=>any;
        extraButtons?: ICommandBarItemProps[];
    }

    interface IAboutUsComplexDataProps extends Omit<IAboutUsComplexDataCommandBarProps, "extraButtons"> {
        properties?: IAboutUsAppWebPartProps;
        value: TAboutUsComplexData;
        showEditControls?: boolean;
        onOrderChange?: (oldIndex: number, newIndex: number)=>any; // returns 
        extraButtons?: ICommandBarItemProps[] | ((key: number, value: TAboutUsComplexData)=>ICommandBarItemProps[]);
    }

    export interface IAboutUsComplexDataDisplayProps extends Omit<IAboutUsComplexDataProps, "value" | "itemIndex"> {
        values: [];
    }
    //#endregion

    //#region COMMAND BAR WITH EDIT & DELETE COMPLEX DATA ITEM BUTTONS
    class ComplexDataCommandBar extends React.Component<IAboutUsComplexDataCommandBarProps> {
        public render(): React.ReactElement<IAboutUsComplexDataCommandBarProps> {
            const key = this.props.itemIndex;
            let commandBarItems: ICommandBarItemProps[] = [
                    {
                        key: `btnEditTask${key}`,
                        text: "Edit",
                        iconProps: { iconName: "Edit" },
                        iconOnly: true,
                        ariaLabel: "Edit this entry.",
                        buttonStyles: { root: { "border": "1px solid", "border-radius": "3px;" }},
                        onClick: (evt)=>{ if (this.props.onEdit) this.props.onEdit(key); }
                    }
                ],
                commandBarFarItems: ICommandBarItemProps[] = [
                    {
                        key: `btnDeleteTask${key}`,
                        text: "Delete",
                        iconProps: { iconName: "Delete" },
                        iconOnly: true,
                        ariaLabel: "Delete this entry.",
                        buttonStyles: { root: { "border": "1px solid", "border-radius": "3px;" }},
                        onClick: (evt)=>{ if (this.props.onDelete) this.props.onDelete(key); }
                    }
                ];
            
            // add extra buttons
            if (this.props.extraButtons && this.props.extraButtons.length > 0) commandBarItems = commandBarItems.concat(this.props.extraButtons);
            
            return (
                <CommandBar
                    items={ commandBarItems }
                    farItems={ commandBarFarItems }
                    className={ styles.aboutUsDisplayItemCommandBar } />
            );
        }
    }
    //#endregion

    //#region TASKS DISPLAY
    class TaskItem extends React.Component<IAboutUsComplexDataProps> {
        public render(): React.ReactElement<IAboutUsComplexDataProps> {
            if (!this.props.value || !this.props.value.text) return null;
            
            const key = this.props.itemIndex,
                value = this.props.value,
                itemClasses = [styles.aboutUsTaskItem],
                tooltipText = [],
                commandBarProps: IAboutUsComplexDataCommandBarProps = {
                    itemIndex: this.props.itemIndex,
                    onEdit: this.props.onEdit,
                    onDelete: this.props.onDelete,
                    extraButtons: (typeof this.props.extraButtons === "function") ? this.props.extraButtons(key, value) : this.props.extraButtons
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            if (value.tooltip) tooltipText.push(value.tooltip);
            if (value.auth) tooltipText.push("Tasking authority: " + value.auth);

            // for SortableItem elements the class names must be global.
            return (
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>

                    <div className={ itemClasses.join(" ") }>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) } >
                            <div className={ styles.task }>{ value.text }</div>
                        </TooltipHost>
                        {
                            (this.props.showEditControls) ? React.createElement(ComplexDataCommandBar, commandBarProps) : null
                        }
                    </div>
                </Wrapper>
            );
        }
    }

    export class TasksDisplay extends React.Component<IAboutUsComplexDataDisplayProps> {
        public render(): React.ReactElement<IAboutUsComplexDataDisplayProps> {
            return (
                <Wrapper
                    condition={ this.props.showEditControls }
                    wrapper={ children => <SortableList
                        allowDrag={ this.props.showEditControls }
                        onSortEnd={ this.props.onOrderChange }
                        className={ (this.props.showEditControls) ? styles.aboutUsSortableList : null }
                        draggedItemClassName={ (this.props.showEditControls) ? styles.aboutUsSortableItemDragged : null } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(TaskItem, {
                                value: value,
                                itemIndex: ndx,
                                showEditControls: this.props.showEditControls,
                                onEdit: this.props.onEdit,
                                onDelete: this.props.onDelete,
                                extraButtons: this.props.extraButtons
                            })
                        )) : null
                    }
                </Wrapper>
            );
        }
    }
    //#endregion

    //#region BIOS DISPLAY
    class BioItem extends React.Component<IAboutUsComplexDataProps> {
        private no_bio_pic = require("./assets/bio_nopic.png");

        public render(): React.ReactElement<IAboutUsComplexDataProps> {
            if (!this.props.value || !this.props.value.name) return null;
            
            const key = this.props.itemIndex,
                value = this.props.value,
                itemClasses = [styles.aboutUsBioItem],
                tooltipText = [],
                commandBarProps: IAboutUsComplexDataCommandBarProps = {
                    itemIndex: this.props.itemIndex,
                    onEdit: this.props.onEdit,
                    onDelete: this.props.onDelete,
                    extraButtons: (typeof this.props.extraButtons === "function") ? this.props.extraButtons(key, value) : this.props.extraButtons
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            if (value.tooltip) tooltipText.push(value.tooltip);

            // for SortableItem elements the class names must be global.
            return (
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>
                
                    <div className={ itemClasses.join(" ") }>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) } >
                            { (value.position) ? 
                                <h3 className={ styles.bioPosition }>{ value.position }</h3>
                                : null
                            }
                            <div className={ styles.bioImgContainer }>
                                <img className={ styles.bioImg } alt="Bio image" src={ value.image || this.no_bio_pic } />
                            </div>
                            <div className={ styles.bioNameContainer } >
                                { (value.bio) ?
                                    <a className={ styles.bioLink } href={ value.bio } target="_blank">
                                        <span className={ styles.bioName }>{ value.name }</span>
                                    </a>
                                    : <span className={ styles.bioName }>{ value.name }</span>
                                }
                            </div>
                            { (value.subtitle) ?
                                <div className={ styles.bioSubtitle }>{ value.subtitle }</div>
                                : null
                            }
                        </TooltipHost>
                        {
                            (this.props.showEditControls) ? React.createElement(ComplexDataCommandBar, commandBarProps) : null
                        }
                    </div>
                </Wrapper>
            );
        }
    }

    export class BiosDisplay extends React.Component<IAboutUsComplexDataDisplayProps> {
        public render(): React.ReactElement<IAboutUsComplexDataDisplayProps> {
            return (
                <Wrapper
                    condition={ this.props.showEditControls }
                    wrapper={ children => <SortableList
                        allowDrag={ this.props.showEditControls }
                        onSortEnd={ this.props.onOrderChange }
                        className={ (this.props.showEditControls) ? styles.aboutUsSortableList : null }
                        draggedItemClassName={ (this.props.showEditControls) ? styles.aboutUsSortableItemDragged : null } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(BioItem, {
                                    value: value,
                                    itemIndex: ndx,
                                    showEditControls: this.props.showEditControls,
                                    onEdit: this.props.onEdit,
                                    onDelete: this.props.onDelete,
                                    extraButtons: this.props.extraButtons
                            })
                        )) : null
                    }
                </Wrapper>
            );
        }
    }
    //#endregion

    //#region LINKS
    class LinkItem extends React.Component<IAboutUsComplexDataProps> {
        public render(): React.ReactElement<IAboutUsComplexDataProps> {
            if (!this.props.value || !this.props.value.url) return null;

            const key = this.props.itemIndex,
                value = this.props.value,
                itemClasses = [styles.aboutUsLinkItem],
                tooltipText = [],
                commandBarProps: IAboutUsComplexDataCommandBarProps = {
                    itemIndex: this.props.itemIndex,
                    onEdit: this.props.onEdit,
                    onDelete: this.props.onDelete,
                    extraButtons: (typeof this.props.extraButtons === "function") ? this.props.extraButtons(key, value) : this.props.extraButtons
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            if (value.tooltip) tooltipText.push(value.tooltip);

            // for SortableItem elements the class names must be global.
            return (            
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>

                    <div className={ itemClasses.join(" ") }>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) } >
                            <a className={ styles.link } href={ value.url } target={ (value.target) ? "_blank" : "_self" }>
                                <FontIcon iconName="Link" className={ styles.linkIcon } />
                                { value.text || value.url }
                            </a>
                        </TooltipHost>
                        {
                            (this.props.showEditControls) ? React.createElement(ComplexDataCommandBar, commandBarProps) : null
                        }
                    </div>
                </Wrapper>
            );
        }
    }

    export class LinksDisplay extends React.Component<IAboutUsComplexDataDisplayProps> {
        public render(): React.ReactElement<IAboutUsComplexDataDisplayProps> {
            return (
                <Wrapper
                    condition={ this.props.showEditControls }
                    wrapper={ children => <SortableList
                        allowDrag={ this.props.showEditControls }
                        onSortEnd={ this.props.onOrderChange }
                        className={ (this.props.showEditControls) ? styles.aboutUsSortableList : null }
                        draggedItemClassName={ (this.props.showEditControls) ? styles.aboutUsSortableItemDragged : null } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(LinkItem, {
                                value: value,
                                itemIndex: ndx,
                                showEditControls: this.props.showEditControls,
                                onEdit: this.props.onEdit,
                                onDelete: this.props.onDelete,
                                extraButtons: this.props.extraButtons
                            })
                        )) : null
                    }
                </Wrapper>
            );
        }
    }
    //#endregion

    //#region SOP
    class SOPItem extends React.Component<IAboutUsComplexDataProps> {
        public render(): React.ReactElement<IAboutUsComplexDataProps> {
            if (!this.props.value || !this.props.value.url) return null;

            const key = this.props.itemIndex,
                value = this.props.value,
                itemClasses = [styles.aboutUsLinkItem],
                tooltipText = [],
                commandBarProps: IAboutUsComplexDataCommandBarProps = {
                    itemIndex: this.props.itemIndex,
                    onEdit: this.props.onEdit,
                    onDelete: this.props.onDelete,
                    extraButtons: (typeof this.props.extraButtons === "function") ? this.props.extraButtons(key, value) : this.props.extraButtons
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            if (value.tooltip) tooltipText.push(value.tooltip);

            // for SortableItem elements the class names must be global.
            return (            
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>

                    <div className={ itemClasses.join(" ") }>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) } >
                            <a className={ styles.link } href={ value.url } target={ (value.target) ? "_blank" : "_self" }>
                                <FontIcon iconName="Processing" className={ styles.linkIcon } />
                                { value.text || value.url }
                            </a>
                        </TooltipHost>
                        {
                            (this.props.showEditControls) ? React.createElement(ComplexDataCommandBar, commandBarProps) : null
                        }
                    </div>
                </Wrapper>
            );
        }
    }

    export class SOPDisplay extends React.Component<IAboutUsComplexDataDisplayProps> {
        public render(): React.ReactElement<IAboutUsComplexDataDisplayProps> {
            return (
                <Wrapper
                    condition={ this.props.showEditControls }
                    wrapper={ children => <SortableList
                        allowDrag={ this.props.showEditControls }
                        onSortEnd={ this.props.onOrderChange }
                        className={ (this.props.showEditControls) ? styles.aboutUsSortableList : null }
                        draggedItemClassName={ (this.props.showEditControls) ? styles.aboutUsSortableItemDragged : null } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(SOPItem, {
                                value: value,
                                itemIndex: ndx,
                                showEditControls: this.props.showEditControls,
                                onEdit: this.props.onEdit,
                                onDelete: this.props.onDelete,
                                extraButtons: this.props.extraButtons
                            })
                        )) : null
                    }
                </Wrapper>
            );
        }
    }
    //#endregion

    //#region Contact
    class ContactItem extends React.Component<IAboutUsComplexDataProps> {
        public render(): React.ReactElement<IAboutUsComplexDataProps> {
            if (!this.props.value || !this.props.value.text) return null;

            const key = this.props.itemIndex,
                value = this.props.value,
                itemClasses = [styles.aboutUsContactItem],
                tooltipText = [],
                commandBarProps: IAboutUsComplexDataCommandBarProps = {
                    itemIndex: this.props.itemIndex,
                    onEdit: this.props.onEdit,
                    onDelete: this.props.onDelete,
                    extraButtons: (typeof this.props.extraButtons === "function") ? this.props.extraButtons(key, value) : this.props.extraButtons
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            if (value.tooltip) tooltipText.push(value.tooltip);

            // for SortableItem elements the class names must be global.
            return (            
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>

                    <div className={ itemClasses.join(" ") }>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) } >
                            <div className={ styles.nameContainer }>
                                { (value.title) ? <div className={ styles.title }>{ value.title }</div> : null}
                                <div className={ styles.name }>{ value.text }</div>
                            </div>
                            { (value.email) ? <a className={ styles.link } href={`mailto:${value.email}`} target="_blank">{value.email}</a> : null }
                            { (value.email2) ? <span className={ styles.link } >SIPR: {value.email2}</span> : null }
                            { (value.email3) ? <span className={ styles.link } >JWIC: {value.email3}</span> : null }
                            { (value.phone1 || value.phone2 || value.dsn) ?
                                <div className={ styles.phoneContainer }>
                                    { (value.phone1 ) ? <span className={ styles.phone }>&#9742;: {value.phone1}</span> : null }
                                    { (value.phone2 ) ? <span className={ styles.phone }>Mobile: {value.phone2}</span> : null }
                                    { (value.dsn ) ? <span className={ styles.phone }>DSN: {value.dsn}</span> : null }
                                </div>
                                : null
                            }
                            { (value.location) ? <div className={ styles.location }>Location: {value.location}</div> : null }
                            { (value.website) ? <a className={ styles.link } href={value.website} target="_blank">Website</a> : null }
                        </TooltipHost>
                        {
                            (this.props.showEditControls) ? React.createElement(ComplexDataCommandBar, commandBarProps) : null
                        }
                    </div>
                </Wrapper>
            );
        }
    }

    export class ContactsDisplay extends React.Component<IAboutUsComplexDataDisplayProps> {
        public render(): React.ReactElement<IAboutUsComplexDataDisplayProps> {
            return (
                <Wrapper
                    condition={ this.props.showEditControls }
                    wrapper={ children => <SortableList
                        allowDrag={ this.props.showEditControls }
                        onSortEnd={ this.props.onOrderChange }
                        className={ (this.props.showEditControls) ? styles.aboutUsSortableList : null }
                        draggedItemClassName={ (this.props.showEditControls) ? styles.aboutUsSortableItemDragged : null } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(ContactItem, {
                                value: value,
                                itemIndex: ndx,
                                showEditControls: this.props.showEditControls,
                                onEdit: this.props.onEdit,
                                onDelete: this.props.onDelete,
                                extraButtons: this.props.extraButtons
                            })
                        )) : null
                    }
                </Wrapper>
            );
        }
    }
    //#endregion
//#endregion

//#region TAGS & KEYWORDS
    //#region INTERFACE
    interface IAboutUsKeywordItemProps {
        properties?: IAboutUsAppWebPartProps;
        itemIndex: number;
        value: string;
        showEditControls?: boolean;
        onDelete?: (index: number)=>any;
        onOrderChange?: (oldIndex: number, newIndex: number)=>any; // returns 
    }

    export interface IAboutUsKeywordsDisplayProps extends Omit<IAboutUsKeywordItemProps, "value" | "itemIndex"> {
        values: string[];
    }
    //#endregion
    
    //#region KEYWORDS DISPLAY
    class KeywordItem extends React.Component<IAboutUsKeywordItemProps> {
        public render(): React.ReactElement<IAboutUsKeywordItemProps> {
            if ((typeof this.props.value !== "string") || trim(this.props.value).length === 0) return null;

            const key = this.props.itemIndex,
                value = trim(this.props.value),
                itemClasses = [styles.aboutUsKeywordItem],
                deleteButtonProps: IButtonProps = {
                    iconProps: { iconName: "Cancel" },
                    disabled: (this.props.showEditControls !== true),
                    className: styles.button,
                    onClick: ()=>{ this.props.onDelete(this.props.itemIndex); }
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            // for SortableItem elements the class names must be global.
            return (            
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>

                    <div className={ itemClasses.join(" ") }>
                            <div className={ styles.keyword }>{value}{
                                (this.props.showEditControls) ? 
                                    <TooltipHost content="Remove">
                                        { React.createElement(IconButton, deleteButtonProps) }
                                    </TooltipHost>
                                    : null
                            }</div>
                    </div>
                </Wrapper>
            );
        }
    }

    export class KeywordsDisplay extends React.Component<IAboutUsKeywordsDisplayProps> {
        public render(): React.ReactElement<IAboutUsKeywordsDisplayProps> {
            return (
                <Wrapper
                    condition={ this.props.showEditControls }
                    wrapper={ children => <SortableList
                        allowDrag={ this.props.showEditControls }
                        onSortEnd={ this.props.onOrderChange }
                        className={ (this.props.showEditControls) ? styles.aboutUsSortableList : null }
                        draggedItemClassName={ (this.props.showEditControls) ? styles.aboutUsSortableItemDragged : null } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(KeywordItem, {
                                value: value,
                                itemIndex: ndx,
                                showEditControls: this.props.showEditControls,
                                onDelete: this.props.onDelete
                            })
                        )) : null
                    }
                </Wrapper>
            );
        }
    }
    //#endregion
//#endregion

//#region GLOBAL DISPLAY HELPERS
// inspiried by array-move (https://github.com/sindresorhus/array-move) by Sindre Sorhus. 25 Oct 2021
export function rearrangeArray(arr: any[], oldIndex: number, newIndex: number) {
    arr = [...arr];
    const startIndex = (oldIndex < 0) ? arr.length + newIndex : oldIndex;

    if (startIndex >= 0 && startIndex < arr.length) {
        const endIndex = (newIndex < 0) ? arr.length + newIndex : newIndex,
            [item] = arr.splice(oldIndex, 1);

        arr.splice(endIndex, 0, item);
    }

    return arr;
}

export function TooltipProps(tooltip: string): ITooltipProps {
    const text = trim(tooltip);
    return {
        onRenderContent: ()=>{
            return (text.length > 0) ? <div style={{whiteSpace: "pre-line"}}>{text}</div> : null;
        }
    };
}
//#endregion


//#region PAGE DISPLAY
export default interface IPageDisplayProps {
    ctx: WebPartContext;
    properties: IAboutUsAppWebPartProps;
    list: DataFactory;
    itemId: number;
    history: History;
    pageLayout?: "default";
}

export default class PageDisplay extends React.Component<IPageDisplayProps, Record<string, any>> {
    //#region PROPERTIES
    private structure: Record<(number | string), IDataStructureItem> = {};
    //#endregion

    //#region RENDER
    constructor(props) {
        super(props);

        this.state = {};
    }

    public render(): React.ReactElement<IPageDisplayProps> {
        return (
            <div className={ styles.defaultPageLayout }>
                { this.displayAppMessaage(this.props.properties.appMessage, this.props.properties.appMessageIsAlert) }

                <div className={ styles.headerSection }>
                    { this.displayLogo(this.state.Logo) }
                    <div className={ styles.menuSection }></div>
                    <div className={ styles.toolbarSection }></div>
                    <div className={ styles.searchSection }></div>
                    <div className={ styles.header }>
                        <h2 className={ styles.headerText }>{ this.state.Title || "" } - { this.state.Name || "" }</h2>
                        <div className={ styles.subtitle }>{ this.state.Description || "" }</div>
                    </div>
                </div>


                <div className={ styles.bodySection }>
                    { this.displayMission(this.state.Mission) }
                    { this.displayTasks(this.state.Tasks) }
                    { this.displayContent(this.state.Content) }
                    { this.displaySubContent(this.state.SubContent) }
                    { this.displayLinks(this.state.Links) }
                    { this.displayContacts(this.state.Contacts) }
                </div>

                <div className={ styles.sideSection }>
                    { this.displayBios(this.state.Bios) }
                    { this.displaySOP(this.state.SOP) }
                    { this.displayOfficeInfoBlock(
                        this.state.Location,
                        this.state.Address,
                        this.state.Phone,
                        this.state.DSN,
                        this.state.FAX,
                        this.state.SignatureBlock
                    ) }
                </div>

                <div className={ styles.footerSection }>
                    <div className={ styles.contentManagersContainer }></div>
                    <div className={ styles.validatedContainer }></div>
                </div>
            </div>

        );
    }

    public async componentDidMount() {
        // get item data & nav items
        const getItem = this.props.list.getItemById_expandFields(this.props.itemId),
            getStructure = this.props.list.getDataStructure(this.props.properties),
            responses = await Promise.all([getItem, getStructure]),
            item = responses[0];

        this.structure = responses[1];
        if (item.ID in this.structure) this.structure[item.ID].data = item;

        LOG("item", item);
        LOG("structure:", this.structure);

        // state
        const state = {};
        for (const key in item) {
            if (key.indexOf("odata") === 0) continue;
            if (Object.prototype.hasOwnProperty.call(item, key)) {
                const data = item[key];
                state[key] = data;
            }
        }

        this.setState(assign(this.state, state));
    }
    //#endregion

    //#region LAYOUTS
    // private DefaultPageLayout(): React.ReactElement {
        
    // }
    //#endregion

    //#region COMPONENTS
    private displayAppMessaage(message: string, isAlert: boolean = false, className?: string): React.ReactElement {
        const css = [styles.messageSection];
        if (isAlert) css.push(styles.isAlert);
        if (className) css.push(className);

        message = trim(message || "");

        return (message) ? <div className={ css.join(" ") }>{ message }</div> : null;
    }

    private displayLogo(logo: {"Url": string, "Description": string}, className?: string): React.ReactElement {
        const css = [styles.logoSection];
        if (className) css.push(className);

        return (logo) ? <div className={ css.join(" ") }>
            <img className={ styles.pageLogo } src={ trim(logo.Url) } alt="About-Us page logo" />
        </div> : null;
    }

    private displayMission(text: string, className?: string): React.ReactElement {
        const css = [styles.missionContainer];
        if (className) css.push(className);

        text = trim(text);

        return (text) ? <div className={ css.join(" ") }>{ text }</div> : null;
    }

    private displayTasks(text: string, className?: string): React.ReactElement {
        const css = [styles.tasksContainer];
        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values) ? <div className={ css.join(" ") }>{ React.createElement(TasksDisplay, { values: values }) }</div> : null;
    }

    private displayContent(text: string, className?: string): React.ReactElement {
        const css = [styles.contentContainer];
        if (className) css.push(className);

        text = trim(text);

        return (text) ? <div className={ css.join(" ") } dangerouslySetInnerHTML={ {__html: text} }/> : null;
    }

    private displaySubContent(text: string, className?: string): React.ReactElement {
        const css = [styles.subContentContainer];
        if (className) css.push(className);

        text = trim(text);

        return (text) ? <div className={ css.join(" ") } dangerouslySetInnerHTML={ {__html: text} }/> : null;
    }

    private displayLinks(text: string, className?: string): React.ReactElement {
        const css = [styles.linksContainer];
        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values) ? 
            <>
                <div className={ styles.sectionBanner }>Links</div>
                <div className={ css.join(" ") }>{ React.createElement(LinksDisplay, { values: values }) }</div> 
            </>
            : null;
    }

    private displayContacts(text: string, className?: string): React.ReactElement {
        const css = [styles.contactsContainer];
        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values) ? 
            <>
                <div className={ styles.sectionBanner }>Contacts</div>
                <div className={ css.join(" ") }>{ React.createElement(ContactsDisplay, { values: values }) }</div> 
            </>
            : null;
    }

    private displayBios(text: string, className?: string): React.ReactElement {
        const css = [styles.biosContainer];
        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values) ? <div className={ css.join(" ") }>{ React.createElement(BiosDisplay, { values: values }) }</div> : null;
    }

    private displaySOP(text: string, className?: string): React.ReactElement {
        const css = [styles.sopsContainer];
        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values) ? 
            <>
                <div className={ styles.sectionBanner }>SOP</div>
                <div className={ css.join(" ") }>{ React.createElement(SOPDisplay, { values: values }) }</div> 
            </>
            : null;
    }

    private displayOfficeInfoBlock(location: string = "", address: string = "", phone: string = "", dsn: string = "", fax: string = "", sig: string = "") {
        const officeInfo: Record<("label" | "text" | "css"), string>[] = [];

        // generate the office information array. this determines the order
        location = trim(location || "");
        if (location) officeInfo.push({ label: "Office Location:", text: location, css: styles.locationContainer });
        
        address = trim(address || "");
        if (address) officeInfo.push({ label: "Mailing Address:", text: address, css: styles.addressContainer });
        
        phone = trim(phone || "");
        if (phone) officeInfo.push({ label: "Office Phone:", text: phone, css: styles.phoneContainer });
        
        dsn = trim(dsn || "");
        if (dsn) officeInfo.push({ label: "Office DSN:", text: dsn, css: styles.dsnContainer });
        
        fax = trim(fax || "");
        if (fax) officeInfo.push({ label: "Office FAX:", text: fax, css: styles.faxContainer });
        
        sig = trim(sig || "");
        if (sig) officeInfo.push({ label: "Signature Block(s):", text: sig, css: styles.signatureBlockContainer });
        
        return (
            <>
                <div className={ styles.sectionBanner }>Office Information</div>
                <div className={ styles.officeInformationContainer }>
                    { officeInfo.map(info => {
                        return (<div className={ styles.officeInfo }>
                            <div className={ styles.officeLabel }>{ info.label }</div>
                            <div className={ info.css }>{ info.text }</div>
                        </div>);
                    }) }
                </div>
            </>
        );
    }
    //#endregion

    //#region HELPERS

    //#endregion
}
//#endregion


//#region PRIVATE LOG
/** Prints out debug messages. Decorated console.info() or console.error() method.
 * @param args Message or object to view in the console. If message starts with "ERROR", DEBUG will use console.error().
 */
function LOG(...args: any[]) {
    // is an error message, if first argument is a string and contains "error" string.
    const isError = (args.length > 0 && (typeof args[0] === "string")) ? args[0].toLowerCase().indexOf("error") > -1 : false;
    args = ["(About-Us AboutUsDisplay.tsx)"].concat(args);

    if (window && window.console) {
        if (isError && console.error) {
            console.error.apply(null, args);

        } else if (console.info) {
            console.info.apply(null, args);

        }
    }
}
//#endregion