// About-Us data display elements.
import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { assign, find, trim } from 'lodash';
import * as moment from 'moment';

// https://docs.microsoft.com/en-us/javascript/api/office-ui-fabric-react?view=office-ui-fabric-react-latest
import { ActionButton, CommandBar, DirectionalHint,
    FontIcon,
    IButtonProps,
    ICommandBarItemProps,
    ICommandBarStyles,
    IconButton,
    ITooltipProps, 
    TooltipHost} from 'office-ui-fabric-react';

//> npm install react-easy-sort
import SortableList, { SortableItem } from 'react-easy-sort';
import AboutUsAppWebPart, { sourceContainsAny, IAboutUsAppWebPartProps, LOG, DEBUG, DEBUG_NOTRACE } from '../AboutUsAppWebPart';
import { Wrapper } from './AboutUsApp';
import DataFactory, { IDataStructureItem, IUserPermissions } from './DataFactory';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { ISearchResult } from '@pnp/sp/search';
import FormControls, { LoadingSpinner, ShowConfigureWebPart } from './FormControls';


//#region ICON
    export interface IIconProps {
        iconName: string;
        className?: string;
    }
    
    export class Icon extends React.Component<IIconProps> {
        public render(): React.ReactElement<IIconProps> {
            return <FontIcon iconName={ this.props.iconName } className={ this.props.className } />;
        }
    }
//#endregion

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
            const key = this.props.itemIndex,
                commandBarItems: ICommandBarItemProps[] = [
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
            if (this.props.extraButtons && this.props.extraButtons.length > 0) commandBarItems.concat(this.props.extraButtons);
            
            return <CommandBar 
                    items={ commandBarItems }
                    farItems={ commandBarFarItems }
                    className={ styles.aboutUsDisplayItemCommandBar } />;
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
                    extraButtons: (typeof this.props.extraButtons === "function") ? 
                        this.props.extraButtons(key, value) : this.props.extraButtons
                };

            if (this.props.showEditControls) itemClasses.push(styles.aboutUsSortableItem);

            if (value.tooltip) tooltipText.push(value.tooltip);
            if (!this.props.properties.showTaskAuth && value.auth) tooltipText.push("Tasking authority: " + value.auth);

            // for SortableItem elements the class names must be global.
            return (
                <Wrapper
                    condition={this.props.showEditControls}
                    wrapper={ children => <SortableItem key={key}>{ children }</SortableItem>}>

                    <div className={ itemClasses.join(" ") }>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) } >
                            <div className={ styles.task }>{ value.text }</div>
                            { (this.props.properties.showTaskAuth && value.auth) ?  
                                <div className={ styles.taskAuthContainer }>
                                    <Icon iconName="Childof" className={ styles.fabricUIIcon }/>
                                    <span className={ styles.taskAuthText }>{ value.auth }</span>
                                </div>
                            : null }
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
                        draggedItemClassName={ styles.aboutUsSortableItemDragged } >{ children }</SortableList> }
                    >
                    {
                        (this.props.values && this.props.values instanceof Array) ?
                            this.props.values.map((value, ndx) => ( React.createElement(TaskItem, {
                                properties: this.props.properties,
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
                    extraButtons: (typeof this.props.extraButtons === "function") ? 
                        this.props.extraButtons(key, value) : this.props.extraButtons
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
                        draggedItemClassName={ styles.aboutUsSortableItemDragged } >{ children }</SortableList> }
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
                    extraButtons: (typeof this.props.extraButtons === "function") ? 
                        this.props.extraButtons(key, value) : this.props.extraButtons
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
                                <Icon iconName="Link12" className={ styles.fabricUIIcon } />
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
                        draggedItemClassName={ styles.aboutUsSortableItemDragged } >{ children }</SortableList> }
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
                    extraButtons: (typeof this.props.extraButtons === "function") ? 
                        this.props.extraButtons(key, value) : this.props.extraButtons
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
                                <Icon iconName="Processing" className={ styles.fabricUIIcon } />
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
                        draggedItemClassName={ styles.aboutUsSortableItemDragged } >{ children }</SortableList> }
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
                    extraButtons: (typeof this.props.extraButtons === "function") ? 
                        this.props.extraButtons(key, value) : this.props.extraButtons
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
                            { (value.email) ? 
                                <a className={ styles.link } href={`mailto:${value.email}`} target="_blank">{value.email}</a> : null }
                            { (value.email2) ? <div className={ styles.redContactsText }>SIPR: {value.email2}</div> : null }
                            { (value.email3) ? <div >JWIC: {value.email3}</div> : null }
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
                        draggedItemClassName={ styles.aboutUsSortableItemDragged } >{ children }</SortableList> }
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
                            <div className={ styles.keyword }>
                                <Icon iconName="TagSolid" className={ styles.fabricUIIcon } />
                                {value}
                                {
                                (this.props.showEditControls) ? 
                                    <TooltipHost content="Remove">
                                        { React.createElement(IconButton, deleteButtonProps) }
                                    </TooltipHost>
                                    : null
                                }
                            </div>
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
                        draggedItemClassName={ styles.aboutUsSortableItemDragged } >{ children }</SortableList> }
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

//#region BREADCRUMB
    //#region INTERFACE
    interface IBreakcrumbItemProps {
        properties: IAboutUsAppWebPartProps;
        item: IDataStructureItem;
        text: string;
        displayType: string;
        onClick: (id: number)=>void;
        className?: string;
        displayAll?: boolean;
    }

    interface IBreadcrumbSubmenuProps {
        properties: IAboutUsAppWebPartProps;
        items: IDataStructureItem[];
        className?: string;
        displayType: string;
        onClick: (id: number)=>void;
        ignoreItems?: (number | string)[];
    }

    export interface IBreadcrumbDisplayProps {
        properties: IAboutUsAppWebPartProps;
        structure: Record<any, IDataStructureItem>;
        displayType: string;
        itemID?: number | string;
        onClick: (id: number)=>void;
    }
    //#endregion

    //#region DISPLAY
    class breadcrumbItem extends React.Component<IBreakcrumbItemProps> {
        public render(): React.ReactElement<IBreakcrumbItemProps> {
            const css = [ styles.breadcrumbItem ],
                isLink = true, //sourceContainsAny(this.props.item.DisplayType, ["_orgType"]),
                url = new URL(location.href || ""),
                urlParams = url.searchParams,
                _onClick = (evt) => {
                    evt.preventDefault();
                    
                    if (this.props.item.DisplayType.indexOf("_orgType") > -1) {
                        const id = (this.props.item.children && this.props.item.children.length > 0) ? this.props.item.children[0].ID : null;
                        if (id) this.props.onClick(id);
                    } else {
                        this.props.onClick(this.props.item.ID);
                    }
                    return false;
                };

            if (this.props.className) css.push(this.props.className);
            urlParams.set(this.props.properties.urlParam, this.props.item.ID.toString());

            return <li className={ css.join(" ") }>
                <Wrapper
                    condition={ isLink }
                    wrapper={ children => <a href={ url.toString() } className={ styles.breadcrumbText } onClick={ _onClick } data-interception="off">{children}</a> }
                    else={ children => <div className={ styles.breadcrumbText }>{children}</div> }
                >
                    { this.props.text }
                </Wrapper>
                { this.props.children }
            </li>;
        }
    }

    class breadcrumbSubmenu extends React.Component<IBreadcrumbSubmenuProps> {
        public render(): React.ReactElement<IBreadcrumbSubmenuProps> {
            const items = this.generateChildrenList(),
                classNames = ["breadcrumbSubmenu", styles.subBreadcrumbContainer];

            return (items.length > 0) ?
                <div className={ classNames.join(" ") }>
                    <ul className={ styles.subBreadcrumbList }>
                        {
                            items.map(item => React.createElement(breadcrumbItem, {
                                properties: this.props.properties,
                                item: item,
                                text: item.Title + ((item.Name) ? " - " + item.Name : ""),
                                displayType: this.props.displayType,
                                onClick: this.props.onClick,
                                className: styles.subBreadcrumbItem })
                            )
                        }
                    </ul>
                </div> 
            : null ;
        }

        private generateChildrenList(): IDataStructureItem[] {
            const allowedTypes = ["_orgType"],
                ignoreList = (this.props.ignoreItems) ? this.props.ignoreItems : [];
            let _items = [];

            this.props.items.forEach(item => {
                // does item contain any of the allowed types && not on the ignore list
                //if (sourceContainsAny(item.DisplayType, allowedTypes) && ignoreList.indexOf(item.ID) === -1) {
                    if (item.DisplayType.indexOf("_orgType") > -1) {
                        // this item is an OrgType (HQ, Task Force, Component, LNO, Other...)
                        // don't show item if it doesn't have a child
                        if (!item.children || item.children.length === 0) {
                            return;
                        }
                    }
                    _items.push(item);
                //}
            });

            // sort
            _items.sort((a, b) => { return a.OrderBy - b.OrderBy; });

            return _items;
        }
        
    }

    export class breadcrumbDisplay extends React.Component<IBreadcrumbDisplayProps> {
        public render(): React.ReactElement<IBreadcrumbDisplayProps> {
            const topItems = this.generateTopBreadcrumbItems(this.props.structure, this.props.itemID);

            return (
                <div className={ styles.breadcrumbSection } >
                    { (topItems.length > 0) ?
                        <ul className={ styles.topBreadcrumbList }>
                            {
                                topItems.map((item, i) => {
                                    const nextItem = (i < topItems.length - 1) ? topItems[i + 1] : null,
                                    classNames = [styles.topBreadcrumbItem],
                                    hasChildren = this.hasChildren(item.children, nextItem, this.props.displayType);

                                    if (hasChildren) classNames.push(styles.hasSubmenu);

                                    return React.createElement(breadcrumbItem,
                                        {
                                            properties: this.props.properties,
                                            item: item,
                                            text: item.Title,
                                            displayType: this.props.displayType,
                                            className: classNames.join(" "),
                                            displayAll: true,
                                            onClick: this.props.onClick
                                        },
                                        React.createElement(breadcrumbSubmenu, {
                                            properties: this.props.properties,
                                            ignoreItems: (nextItem && nextItem.ID !== 0) ? [nextItem.ID] : null,
                                            items: item.children,
                                            displayType: this.props.displayType,
                                            onClick: this.props.onClick,
                                            className: styles.subBreadcrumbItem }
                                        )
                                    );
                                })
                            }
                        </ul>
                    : null }
                </div>
            );
        }

        private generateTopBreadcrumbItems(
            structure: Record<(number | string),
            IDataStructureItem>, itemID: number | string): IDataStructureItem[] {

            const items = [];

            // keep getting the parent item until null or "_root"
            let item = structure[itemID];
            if (item) items.push(item);
            while (item) {
                item = structure[item.ParentID];
                if (item) items.unshift(item);
            }

            return items;
        }

        private hasChildren(children: IDataStructureItem[], nextItem: IDataStructureItem, displayType: string): boolean {
            const allowedTypes = [displayType, "_orgType"],
                ignoreList = (nextItem && nextItem.ID !== 0) ? [nextItem.ID] : [];

            for (let i = 0; i < children.length; i++) {
                const item = children[i];
                
                // does item contain any of the allowed types && not on the ignore list
                //if (sourceContainsAny(item.DisplayType, allowedTypes) && ignoreList.indexOf(item.ID) === -1) {
                    if (item.DisplayType.indexOf("_orgType") > -1) {
                        // this item is an OrgType (HQ, Task Force, Component, LNO, Other...)
                        // show item if it has a child
                        if (item.children && item.children.length > 0) {
                            return true;
                        }
                    } else {
                        return true;
                    }
                //}
            }

            return false;
        }
    }
    //#endregion
//#endregion

//#region CONTENT MANAGERS
    export interface IContentManagersDisplayProps {
        users: IUserInfo[];
        ownerGroupID: number;
        emailSubject: string;
    }
    interface IContentManagersDisplayState {
        owners: ISiteUserInfo[];
        emails: string[];
    }
    
    export class ContentManagersDisplay extends React.Component<IContentManagersDisplayProps, IContentManagersDisplayState> {
        constructor(props) {
            super(props);

            this.state = {
                owners: [],
                emails: []
            };
        }

        public render(): React.ReactElement<IContentManagersDisplayProps> {
            const cmMailToLink = this.generateContentManagersMailToLink();

            // if content managers mailto link === null; this means there are no users or owners
            return (cmMailToLink) ? 
                <>
                    { (this.props.users && this.props.users.length > 0) ?
                        <ul className={ styles.contentManagersList }>
                            {(this.state.emails.length > 0) ? this.props.users.map(user => <li>{this.mailTo(user)}</li>) : null }
                        </ul> : null
                    }
                    <div className={ styles.contentManagersMessage }>
                        Have questions, corrections or comments about this page, send a message to the <a href={cmMailToLink}>Content Managers</a>.
                    </div>
                </> : null;
        }

        public async componentDidMount() {
            let owners = [],
                emails = [];

            if (this.props.ownerGroupID) {
                owners = await DataFactory.getSiteGroupMembers(this.props.ownerGroupID);
            }

            if (this.props.users && this.props.users.length > 0) {
                emails = await this.getContentManagersEmails(this.props.users);
            }

            this.setState({
                "owners": owners,
                "emails": emails
            });
        }

        private async getContentManagersEmails(users: IUserInfo[]): Promise<string[]> {
            const emails = [],
                promises = (users && users.length) 
                ? users.map(user => DataFactory.getUserById(user.ID))
                : null ;

            if (promises) {
                const responses = await Promise.all(promises);

                responses.forEach(userInfo => {
                    if (userInfo.Email) {
                        const user = find(users, {"ID": userInfo.Id});
                        if (user) user.EMail = userInfo.Email;
                        emails.push(userInfo.Email);
                    }
                });
            }

            return emails;
        }

        private generateContentManagersMailToLink(): string {
            const users = [],
                owners = [];

            let subject = encodeURIComponent(this.props.emailSubject || "About-Us"),
                body = encodeURIComponent(location.href) || "",
                mailTo = "";

            // add user emails to list
            if (this.props.users && this.props.users.length > 0) {
                this.props.users.forEach(user => {
                   if (user.EMail) users.push(user.EMail);
                });
            }

            // add owner emails to list
            if (this.state.owners && this.state.owners.length > 0) {
                this.state.owners.forEach(owner => {
                    if (owner.Email) owners.push(owner.Email);
                });
            }

            // compile all parts of the mailto link
            const _createLink = () => {
                var to = (users.length > 0) ? users.join(";") : (owners.length > 0) ? owners.join(";") : null,
                    params = [];

                if (users.length > 0 && owners.length > 0) params.push("cc=" + owners.join(";"));
                if (subject) params.push("subject=" + subject);
                if (body) params.push("body=" + body);

                return (to) ? "mailto:" + to + ((params.length > 0) ? "?" + params.join("&"): "") : "";
            };

            // reduce the mailto link parts smartly
            const _reduceLink = () => {
                
                if (owners.length > 2) {
                    // 1. reduce owner emails
                    owners.pop();

                } else if (users.length > 2) {
                    // 2. reduce user emails
                    users.pop();

                } else if (body.length > 0) {
                    // 3. remove body text
                    body = "";

                } else {
                    // 4. remove subject text
                    subject = "";
                }

                return _createLink();
            };

            // ensure mailto link is not longer than 1900 characters
            mailTo = _createLink();
            while (mailTo.length > 1900) {
                mailTo = _reduceLink();
            }

            return mailTo;
        }

        private mailTo(user: IUserInfo): React.ReactElement {
            return <Wrapper
                    condition={ "EMail" in user && user.EMail.length > 0 }
                    wrapper={ children => 
                        <a href={ `mailtto:${user.EMail}?subject=${encodeURIComponent(this.props.emailSubject)}` } className={ styles.email }>
                            {children}
                        </a> }
                    else={ children => <span className={ styles.email }>{children}</span>}
                >
                    <Icon iconName="Mail" className={ styles.fabricUIIcon } />
                    {user.Title || user.EMail || user.Name || "Content Manager"}
                </Wrapper>;
        } 
    }
//#endregion

//#region PAGE VALIDATION
export interface IPageValidationDisplayProps {
    properties: IAboutUsAppWebPartProps;
    onValidate?: ()=>void;
    showButton?: boolean;
    validated: Date;
    validatedBy: IUserInfo;
}

export class PageValidationDisplay extends React.Component<IPageValidationDisplayProps> {
    private friendlyDateFormat = "M/D/YYYY h:mm A";

    public render(): React.ReactElement<IPageValidationDisplayProps> {
        const expirationDate = this.getExpirationDate(this.props.validated),
            expired = this.expired(this.props.validated),
            showWarning = this.showWarning(this.props.validated),
            isValid = this.isValid(this.props.validated),
            expiresIn = (!expired && expirationDate) ? expirationDate.diff(moment(), "days") : 0,
            tooltipText = [],
            css = [styles.validateButton],
            btnProps: IButtonProps = {
                text: (this.props.validated) ? "Update validation status" : "Validate site information",
                iconProps: { iconName: "CompletedSolid" },
                className: "",
                onClick: (evt) => {this.props.onValidate(); return false;}
            };

        if (expirationDate) {
            if (expired ) {
                css.push(styles.validatedExpired);
                tooltipText.push(`Expired! Content Manager(s) are required to validate page information every \
                    ${this.props.properties.validateEvery.toString()} days.\
                    \nPage validation expired on ${expirationDate.format(this.friendlyDateFormat)}.`);

            } else if (showWarning) {
                css.push(styles.validatedWarning);
                tooltipText.push(`Page validation expires soon. Content Manager(s) \
                    should validate page information before the expiration date.\
                    \nExpires on ${expirationDate.format(this.friendlyDateFormat)} \
                    (${expiresIn} day${(expiresIn === 1) ? "" : "s"}).`);
                
            } else if (isValid) {
                css.push(styles.validatedGood);
                tooltipText.push(`Page validated. 
                    ${ (this.props.properties.validateEvery !== 0) ? `Content Manager(s) are required to validate page information every \
                        ${this.props.properties.validateEvery.toString()} days.`: `` } \
                    \nExpires on ${expirationDate.format(this.friendlyDateFormat)} \
                    (${expiresIn} day${(expiresIn === 1) ? "" : "s"}).`);
                
            }
        }
        btnProps.className = css.join(" ");

        return (this.props.properties.validateEvery > 0) ?
            <>
                { (moment.isDate(this.props.validated)) ? <div className={ styles.validatedText }>
                    {
                        `Information on this page was validated \
                        by ${this.props.validatedBy.Title || this.props.validatedBy.EMail || this.props.validatedBy.Name || "unknown"} \
                        on ${moment(this.props.validated).format(this.friendlyDateFormat)}.`
                    }
                </div> : null}
                { (this.props.showButton) ? 
                    <div>
                        <TooltipHost tooltipProps={ TooltipProps(tooltipText.join("\n")) }>
                            { React.createElement(ActionButton, btnProps) }
                        </TooltipHost>
                    </div>
                 : null }
            </>
        : null;
    }

    /** Gets the expiration date based on the given date + the web part property date
     * @param date Starting date object or string.
     * @returns New Date (moment) object that represents the expiration date.
     */
    private getExpirationDate(date: Date | moment.Moment | string): moment.Moment {
        if (moment.isDate(date) && this.props.properties.validateEvery > 0) {
            return moment(date).add(this.props.properties.validateEvery + 1, "days");
        }

        return null;
    }

    /** Check to see if today is after the expiration date
     * @param date Starting date object.
     * @returns True of today is after the expiration date.
     */
    private expired(date: Date): boolean {
        // never expires if validate every 0 days
        if (this.props.properties.validateEvery === 0) return false;

        if (moment.isDate(date)) {
            const today = moment(),
                expirationDate = this.getExpirationDate(date);

            return today.isAfter(expirationDate);
        }

        return true;
    }

    /** Check to see if today is 'near' the expiration date
     * @param date Starting date object.
     * @returns True of today is 'near' the expiration date.
     */
    private showWarning(date: Date): boolean {
        // never expires if validate every 0 days
        if (this.props.properties.validateEvery === 0) return false;

        if (moment.isDate(date)) {
            const today = moment(),
                warningPeriod = Math.ceil(this.props.properties.validateEvery * .1),    // 10%
                warningDate = moment(date).add((this.props.properties.validateEvery + 1) - warningPeriod, "days");

            return today.isSameOrAfter(warningDate);
        }

        return true;
    }

    /** Check to see if today is before the expiration date
     * @param date Starting date object.
     * @returns True of today is before the expiration date.
     */
    private isValid(date: Date): boolean {
        // never expires if validate every 0 days: neither good or bad
        if (this.props.properties.validateEvery === 0) return false;

        if (moment.isDate(date)) {
            const today = moment(),
                expirationDate = this.getExpirationDate(date);

            return today.isSameOrBefore(expirationDate);
        }

        return false;
    }
}
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
        directionalHint: DirectionalHint.topLeftEdge,
        onRenderContent: ()=>{
            return (text.length > 0) ? <div style={{whiteSpace: "pre-line"}}>{text}</div> : null;
        }
    };
}
//#endregion

//#region SEARCHBOX
export interface ISearchBoxProps {
    properties: IAboutUsAppWebPartProps;
    list: DataFactory;
    structure: TStructure;
    onResultsClick: (id: number) => void;

    className?: string;
    placeholder?: string;
    icon?: string;
}
export interface ISearchBoxState {
    queryText: string;
    showResults: boolean;
    results: ISearchResult[];
    searching: boolean;
    searchboxFocus: boolean;
    resultsFocus: boolean;
}

export class SearchBox extends React.Component<ISearchBoxProps, ISearchBoxState> {
    constructor(props) {
    	super(props);

    	// set initial state
    	this.state = {
            searching: false,
            queryText: "",
            showResults: false,
            searchboxFocus: false,
            resultsFocus: false,
            results: []
        };
    }

    public render(): React.ReactElement<ISearchBoxProps> {
        const css = [styles.searchContainer],
            buttonProps: IButtonProps = {
                className: styles.button,
                iconProps: { iconName: this.props.icon || "Search" },
                disabled: this.state.searching,
                onClick: this.searchButton_onClick.bind(this)
            };

        if (this.props.className) css.push(this.props.className);

        return (
            <div className={ css.join(" ") }
                onBlur={ evt => this.setFocus("searchbox", false) }
                onFocus={ evt =>this.setFocus("searchbox", true) }>
                <div className={ styles.searchboxWrapper }>
                    <input
                        className={ styles.searchbox }
                        value={ this.state.queryText }
                        placeholder={ this.props.placeholder || "Search..." }
                        onChange={ evt => this.searchQuery_onChange(evt.target.value) }
                        onKeyPress={ evt => {
                            if (evt.keyCode === 13 || evt.which === 13) this.searchButton_onClick();
                            return false;
                        } } />
                    <TooltipHost content="Search About-Us content">
                        { React.createElement(IconButton, buttonProps) }
                    </TooltipHost>
                </div>
                <div className={ styles.searchResultsWrapper }
                    style={{"display": (this.state.showResults) ? "" : "none"}}
                    onMouseLeave={ evt => this.setFocus("results", false) }
                    onMouseEnter={ evt =>this.setFocus("results", true) }>
                        <ul className={ styles.searchResultsList }>
                        {
                            (this.state.results.length > 0)
                            ? this.state.results.map(result => this.SearchResultElement(result))
                            : <li className={ styles.searchResult }>No search results for this query</li>
                        }
                    </ul>
                </div>
            </div>
        );
    }

    private SearchResultElement(result: ISearchResult): React.ReactElement {
        const resultPath = new URL(result.Path),
            id = parseInt(resultPath.searchParams.get("ID")),
            item = (!isNaN(id)) ? this.props.structure[id] || null : null,
            url = new URL(window.location.href),
            _onClick = (evt) => {
                evt.preventDefault();
                this.props.onResultsClick(id);
                this.setState({ "showResults": false });
                return false;
            };

        url.searchParams.set(this.props.properties.urlParam, id.toString());

        return (item) ? <li className={ styles.searchResult }>
                <a href={ url.toString() } onClick={ _onClick } data-interception="off">{ `${item.Title} - ${item.Name}` }</a>
            </li> : null;
    }

    private searchQuery_onChange(text) {
        this.setState({
            "queryText": text,
            "showResults": false
        });
    }

    private async searchButton_onClick(): Promise<void> {
        // clear search results
        this.setState({
            "showResults": false,
            "results": []
        });

        // get search term
        const term = trim(this.state.queryText);
        if (!term || term.length < 2) return;


        // get search results
        const results = await this.props.list.search(this.state.queryText);

        this.setState({
            "results": results,
            "showResults": true
        });
    }

    private setFocus(section: "searchbox" | "results", hasFocus: boolean) {
        const searchboxFocus = (section === "searchbox") ? hasFocus : this.state.searchboxFocus,
            resultsFocus = (section === "results") ? hasFocus : this.state.resultsFocus,
            hideResults = (searchboxFocus === false && resultsFocus === false),
            state = {...this.state, [section + "Focus"]: hasFocus};

        // hide search results if searchbox or search results looses focus
        if (hideResults) state.showResults = false;
        
        // update focus state
        this.setState(state);
    }
}
//#endregion


//#region PAGE DISPLAY
    //#region INTERFACES & TYPES
    export default interface IPageDisplayProps {
        ctx: WebPartContext;
        properties: IAboutUsAppWebPartProps;
        list: DataFactory;
        itemId: number;
        changeDisplay: (displayType: string)=>void;
        changeItem: (id: number, title?: string, url?: string, replaceState?: boolean) => void;
    }
    export interface IPageDisplayState {
        itemId: number;
        permissions: IUserPermissions;
        dataStatus: TDataStatusTypes;
        [key: string]: any;
    }

    export interface IUserInfo {
        "odata.type"?: string;
        "odata.id"?: string;
        "ID": number;
        "Title"?: string;
        "Name"?: string;
        "EMail"?: string;
    }

    export type TDataStatusTypes = "init" | "nodata" | "invalid" | "loading" | "ready";

    export type TStructure = Record<(number | string), IDataStructureItem>;
    
    //#endregion

export default class PageDisplay extends React.Component<IPageDisplayProps, IPageDisplayState> {
    //#region PROPERTIES
    public static readonly type = "About-Us Page";  // must match one of the List's 'DisplayType' field options

    private structure: TStructure = {};
    //#endregion

    //#region RENDER
    constructor(props) {
        super(props);

        this.state = {
            dataStatus: "init",
            itemId: null,
            permissions: {
                canAdd: false,
                canEdit: false,
                canDelete: false
            }
        };
    }

    public render(): React.ReactElement<IPageDisplayProps> {
        // get item data only if different. usually from history state changes (history.pushState or window.onpopstate)
        if (this.state.itemId !== this.props.itemId) this.getItem(this.props.itemId);

        switch (this.state.dataStatus) {
            case "init":    // just starting up
            case "loading": // loading data
                return <LoadingSpinner />;

            case "nodata":  // list is empty or no items returned
                return <ShowConfigureWebPart
                    onConfigure={ () => { this.props.changeDisplay("new");} }
                    iconName="Org"
                    iconText="There are no items to display."
                    description={ `When adding an item, select "${PageDisplay.type}" as the display type.`}
                    buttonLabel="Add New Item"
                />;

            case "invalid": // something went wrong?
                return <div className="ms-error">
                    <Icon iconName="Error" className={ styles.fabricUIIcon }/>
                    Something went wrong. Please contact the site's administrator. 
                    Detailed error messages may have been written to the Debug Console.
                </div>;
        
            case "ready":
                switch (this.props.properties.pageTemplate) {
                    // add more templates here. remember to update the SCSS

                    default:    // "default"
                        return this.defaultDisplayTemplate();
                }

            default:
                return null;
        } 
    }

    public async componentDidMount() {
        // get item data & breadcrumb items
        const [ structure, permissions] = await Promise.all([
                this.props.list.getDataStructure(this.props.properties.homeTitle),
                this.props.list.getUserPermissions()
            ]);

        this.structure = structure;

        this.setState({...this.state, permissions: permissions});
        

        // initialize state with item data
        const item = await this.getItem(this.props.itemId);
    }
    //#endregion

    //#region DISPLAYS/TEMPLATES
    private defaultDisplayTemplate(): React.ReactElement {
        return <div className={ styles.defaultPageLayout }>
            { this.displayAppMessaage(this.props.properties.appMessage, this.props.properties.appMessageIsAlert) }
            { this.displayMenu(false) }

            <div className={ styles.headerSection }>
                { this.displayLogo(this.state.Logo) }
                { this.displayBreadcrumb() }
                { this.displaySearch() }
                { this.displayHeaderTitle(this.state.Title, this.state.Name, this.state.Description) }
            </div>

            { (this.state.DisplayType && this.state.DisplayType.indexOf(PageDisplay.type) > -1) ? 
                <>
                    <div className={ styles.bodySection }>
                        { this.displayMission(this.state.Mission) }
                        { this.displayTasks(this.state.Tasks) }
                        { this.displayContent(this.state.Content) }
                        { this.displaySubContent(this.state.SubContent) }
                        { this.displayKeywords(this.state.Keywords, "", true) }
                        { this.displayLinks(this.state.Links, "", true) }
                        { this.displayContacts(this.state.Contacts, "", true) }
                    </div>

                    <div className={ styles.sideSection }>
                        { this.displayBios(this.state.Bios) }
                        { this.displaySOP(this.state.SOP, "", true) }
                        { this.displayOfficeInfoBlock(
                            this.state.Location,
                            this.state.Address,
                            this.state.Phone,
                            this.state.DSN,
                            this.state.FAX,
                            this.state.SignatureBlock,
                            true
                        ) }
                    </div>

                    <div className={ styles.footerSection }>
                        { this.displayContentManagers(this.state.ContentManagers, "", true) }
                        { this.displayPageValidation(this.state.Validated, this.state.ValidatedBy) }
                    </div>
                </>
                : <div className={ styles.bodySection }>This About-Us page is not availble.</div>
            }
        </div>;
    }
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
        const css = [styles.logoSection],
            imgUrl = (logo) ? trim(logo.Url) : (this.props.properties.logo) ? this.props.properties.logo.fileAbsoluteUrl || "" : "" ;

        if (className) css.push(className);

        return (imgUrl) ? <div className={ css.join(" ") }>
            <img className={ styles.pageLogo } src={ imgUrl } alt="About-Us page logo" />
        </div> : null ;
    }

    private displayMenu(showViews: boolean = true, showTools: boolean = true, className?: string): React.ReactElement {
        const css = [styles.menuSection],
            key = this.props.ctx.webPartTag,
            commandBarStyles: ICommandBarStyles = {
                root: { "fontSize": "12px", "height": "auto" },
                primarySet: { "fontSize": "12px" },
                secondarySet: { "fontSize": "12px" }
            },
            items: ICommandBarItemProps[] = [],
            farItems: ICommandBarItemProps[] = [];

        if (className) css.push(className);

        if (showViews) {
            // 'Org Chart' button
            items.push({
                key: `btnOrgChart${key}`,
                text: "Org Chart",
                iconProps: { iconName: "Org", styles: {root: {"fontSize": "12px"}} },
                className: styles.menuItem,
                onClick: evt => { this.props.changeDisplay("orgchart"); }
            });

            // 'Accordian' button
            items.push({
                key: `btnAccordian${key}`,
                text: "Explorer",
                iconProps: { iconName: "DOM", styles: {root: {"fontSize": "12px"}} },
                className: styles.menuItem,
                onClick: evt => { this.props.changeDisplay("accordian"); }
            });

            // 'Phone' button
            items.push({
                key: `btnPhone${key}`,
                text: "Phone Directory",
                iconProps: { iconName: "PublishCourse", styles: {root: {"fontSize": "12px"}} },
                className: styles.menuItem,
                onClick: evt => { this.props.changeDisplay("phone"); }
            });
        }

        if (showTools) {
            // 'New' button
            if (this.state.permissions.canAdd) farItems.push({
                key: `btnNew${key}`,
                text: "Add Item",
                iconProps: { iconName: "Add", styles: {root: {"fontSize": "12px"}} },
                className: styles.menuItem,
                onClick: evt => { this.props.changeDisplay("new"); }
            });

            // 'Edit' button
            if (this.state.permissions.canEdit && this.state.ID) farItems.push({
                key: `btnEdit${key}`,
                text: "Edit Item",
                iconProps: { iconName: "Edit", styles: {root: {"fontSize": "12px"}} },
                className: styles.menuItem,
                onClick: evt => { this.props.changeDisplay("edit"); }
            });
        }

        return (items.length > 0 || farItems.length > 0)
            ? <CommandBar items={ items } farItems={ farItems } className={ css.join(" ") } styles={ commandBarStyles } />
            : null;
    }
    
    private displayBreadcrumb(): React.ReactElement {
        const props: IBreadcrumbDisplayProps = {
            properties: this.props.properties,
            structure: this.structure,
            displayType: PageDisplay.type,
            itemID: this.props.itemId,
            onClick: this.navigateTo.bind(this)
        };

        return (Object.keys(this.structure).length > 0)
            ? React.createElement(breadcrumbDisplay, props)
            : null;
    }

    private displaySearch(): React.ReactElement {
        const props: ISearchBoxProps = {
            properties: this.props.properties,
            list: this.props.list,
            structure: this.structure,
            onResultsClick: this.navigateTo.bind(this)
        };

        return (Object.keys(this.structure).length > 0) ? <div className={ styles.searchSection }>
            { React.createElement(SearchBox, props) }
        </div> : null;
    }

    private displayHeaderTitle(title: string, name: string, description: string, className?: string): React.ReactElement {
        const css = [styles.header],
            headerTitle = [];

        if (title) headerTitle.push(title);
        if (name) headerTitle.push(name);

        if (className) css.push(className);

        return (headerTitle.length > 0)
            ? <div className={ css.join(" ") }>
                <h2 className={ styles.headerText }>{ headerTitle.join(" - ") }</h2>
                { (description) ? <div className={ styles.subtitle }>{ description || "" }</div> : null}
            </div>
            : null;
    }

    private displayMission(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Mission"),
            css = [styles.missionContainer];

        if (!showBanner) css.push(styles.showLabel);
        if (className) css.push(className);

        text = trim(text);

        return (text)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") } data-label={ field.Title + ":" }>{ text }</div>
            </>
            : null;
    }

    private displayTasks(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Tasks"),
            css = [styles.tasksContainer];

        if (!showBanner) css.push(styles.showLabel);
        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values && values.length > 0 && Object.keys(values[0]).length > 0)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") }  data-label={ field.Title + ":" }>
                    { React.createElement(TasksDisplay, { values: values, properties: this.props.properties }) }
                </div>
            </>
            : null;
    }

    private displayContent(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Content"),
            css = [styles.contentContainer];

        if (className) css.push(className);

        text = trim(text);

        return (text)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") } dangerouslySetInnerHTML={ {__html: text} }/>
            </>
            : null;
    }

    private displaySubContent(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("SubContent"),
            css = [styles.subContentContainer];

        if (className) css.push(className);

        text = trim(text);

        return (text)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") } dangerouslySetInnerHTML={ {__html: text} }/>
            </>
            : null;
    }

    private displayKeywords(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Keywords"),
            css = [styles.keywordsContainer];

        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values && values.length > 0 && (typeof values[0] === "string"))
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") }>{ React.createElement(KeywordsDisplay, { values: values }) }</div> 
            </>
            : null;
    }

    private displayLinks(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Links"),
            css = [styles.linksContainer];

        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values && values.length > 0 && Object.keys(values[0]).length > 0)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") }>{ React.createElement(LinksDisplay, { values: values }) }</div> 
            </>
            : null;
    }

    private displayContacts(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Contacts"),
            css = [styles.contactsContainer];

        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values && values.length > 0 && Object.keys(values[0]).length > 0)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") }>{ React.createElement(ContactsDisplay, { values: values }) }</div> 
            </>
            : null;
    }

    private displayBios(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("Bios"),
            css = [styles.biosContainer];

        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values && values.length > 0 && Object.keys(values[0]).length > 0)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") }>{ React.createElement(BiosDisplay, { values: values }) }</div>
            </>
            : null;
    }

    private displaySOP(text: string, className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("SOP"),
            css = [styles.sopsContainer];

        if (className) css.push(className);

        const values = (text) ? JSON.parse(trim(text)) : null;

        return (values && values.length > 0 && Object.keys(values[0]).length > 0)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                <div className={ css.join(" ") }>{ React.createElement(SOPDisplay, { values: values}) }</div> 
            </>
            : null;
    }

    private displayOfficeInfoText(text: string, className?: string, label?: string): React.ReactElement {
        return (
            <div className={ styles.officeInfo }>
                { (label) ? <div className={ styles.officeLabel }>{ label }</div> : null }
                <div className={ className }>{ text }</div>
            </div>
        );
    }

    private displayOfficeInfoBlock(
        location: string = "",
        address: string = "",
        phone: string = "",
        dsn: string = "",
        fax: string = "",
        sig: string = "",
        showBanner?: boolean): React.ReactElement {

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
        
        return ( (officeInfo.length > 0) ?
            <>
                { (showBanner) ? <div className={ styles.sectionBanner }>Office Information</div> : null }
                <div className={ styles.officeInformationContainer }>
                    { officeInfo.map(info => this.displayOfficeInfoText(info.text, info.css, info.label)) }
                </div>
            </> : null
        );
    }

    private displayContentManagers(users: IUserInfo[], className?: string, showBanner?: boolean): React.ReactElement {
        const field = this.props.list.getExistingField_InternalName("ContentManagers"),
            css = [styles.contentManagersContainer],
            props = {
                users: users || [],
                ownerGroupID: this.props.properties.ownerGroup,
                emailSubject: `About-Us Question: ${this.state.Title}`
            };

        if (className) css.push(className);

        // render if: item data was fetched AND there is either users or an owner
        return (this.state.ID && ((users && users.length > 0) || this.props.properties.ownerGroup))
            ? <>
                    { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                    <div className={css.join(" ") }>{ React.createElement(ContentManagersDisplay, props) }</div>
                </>
            : 
                <div>
                    { (showBanner) ? <div className={ styles.sectionBanner }>{ field.Title }</div> : null }
                    <div className={css.join(" ") }>
                        <Icon iconName="IncidentTriangle" className={ styles.fabricUIIcon }/>
                        <span>
                            There are no Content Managers, Owners, or Admins for this page. 
                            Please contact the site's owner or admins to update permissions to this web part.
                        </span>
                    </div>
                </div>
            ;
    }

    private displayPageValidation(validated: string, validatedBy: IUserInfo, className?: string, showBanner?: boolean): React.ReactElement {
        const css = [styles.validatedContainer],
            props: IPageValidationDisplayProps = {
                properties: this.props.properties,
                validated: (validated) ? new Date(validated) : null,
                validatedBy: validatedBy,
                showButton: this.state.permissions.canEdit,
                onValidate: this.validate_onClick.bind(this)
            };

        if (className) css.push(className);

        return (moment.isDate(validated) || props.showButton)
            ? <>
                { (showBanner) ? <div className={ styles.sectionBanner }>Validate Page Information</div> : null }
                <div className={css.join(" ") }>{ React.createElement(PageValidationDisplay, props) }</div>
            </>
            : null ;
    }
    //#endregion

    //#region HELPERS
    /** Get item data. Resets the state object to remove previous item data from it.
     * - Assumes the Structure object exists and populated.
     * @param id ID of item to fetch
     * @param initState Optional. Initial starting state. Useful if the state was pre-built prior to calling this method.
     * @returns Item data object
     */
    private async getItem(id: number): Promise<any> {
        const structureItem = (id in this.structure) ? this.structure[id] : null ;
        let item = null;

        // create new state without all the item data.
        // need to copy all non-item data values.
        const initState: IPageDisplayState = {
            dataStatus: "loading",
            itemId: id, // update item id value
            permissions: this.state.permissions // copy from original state
        };

        // if structure object is empty. propably because the list is new/empty.
        if (Object.keys(this.structure).length === 0) {
            initState.dataStatus = "nodata";
            this.setState(initState);
            return null;    // exit now, do not continue
        }

        // get item from cache
        if (structureItem && structureItem.data && Object.keys(structureItem).length > 0) {
            item = structureItem.data;
        }

        // try to fetch the item data
        if (!item) {
            try{
                item = await this.props.list.getItemById_expandFields(id);
            } catch (er) {
                item = null;
                LOG(`ERROR! Unable to get item data for ID: ${id}.`, er);
            }
        }

        // if item wasn't fetchable...
        if (!item) {
            // try the startingID (if set) AND current ID isn't the startingID
            if (typeof this.props.properties.startingID === "number" && id !== this.props.properties.startingID && this.props.properties.startingID > 0) {
                this.props.changeItem(this.props.properties.startingID, null, null, true);
                return;

            } else {
                // try showing the first structure item
                id = null;
                for (const key in this.structure) {
                    const _item = this.structure[key],
                        _id = parseInt(key);

                    // key's in the structure object can be literal strings or numbers. Numeric keys are list items.
                    // check to see if key is a number, can be displayed, and wasn't already tried before
                    if (!isNaN(_id) && _item.DisplayType.indexOf(PageDisplay.type) > -1 && !_item.flags.attemptedFetch) {
                        id = _id;
                        _item.flags.attemptedFetch = true;
                        break;
                    }
                }

                if (id) {
                    this.props.changeItem(id, null, null, true);
                    return;

                } else {
                    // there are items in the list but none of them are for this view
                    initState.dataStatus = "nodata";
                    this.setState(initState);
                    return null;    // exit now, do not continue
                }
            }
        }

        // finally, was item found
        //if (item && "DisplayType" in item && item.DisplayType instanceof Array && item.DisplayType.indexOf(PageDisplay.type) > -1) {
        if (item) {
            initState.dataStatus = "ready";

            // keep track of retrieved items
            if (item.ID in this.structure) this.structure[item.ID].data = item;

            // add item data to state
            for (const key in item) {
                if (key.indexOf("odata") === 0) continue;
                if (Object.prototype.hasOwnProperty.call(item, key)) {
                    const data = item[key];
                    initState[key] = data;
                }
            }

        } else {
            // if still no data by this point, all item ID(s) are invalid
            initState.dataStatus = "invalid";
        }

        this.setState(initState);

        return item;
    }

    /** Navigate to a different item
     * @param id Item ID to navigate to.
     */
    private navigateTo(id: number) {
        const item = this.structure[id] || null;

        if (item) {
            this.props.changeItem(id, item.Title);
        }
    }

    /** Validate Page Info button onClick handler */
    private validate_onClick() {
        const today = new Date();

        // get current user
        this.props.list.getCurrentUser().then(user => {
            // update list item
            this.props.list.api.items.getById(this.props.itemId).update({
                "Validated": today.toISOString(),
                "ValidatedById": user.Id
            }).then(() => {
                this.setState({
                    "Validated": today.toISOString(),
                    "ValidatedById": user.Id,
                    "ValidatedBy": {
                        "odata.type": user["odata.type"],
                        "odata.id": user["odata.id"],
                        "ID": user.Id,
                        "Title": user.Title,
                        "Name": user.LoginName,
                        "EMail": user.Email
                    }
                });
            });
        });
    }
    //#endregion
}
//#endregion