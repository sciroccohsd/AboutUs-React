import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { trim, escape } from 'lodash';

import DataFactory from './DataFactory';
import CustomDialog from './CustomDialog';

import FormControls, {IFormControlsProps} from './FormControls';

export interface IAboutUsAppProps {
    displayType: string;
    list: DataFactory;
  }

export default class AboutUsApp extends React.Component<IAboutUsAppProps, {}> {
    //#region PROPERTIES
    private list: DataFactory;
    private displayType: string;
    //#endregion

    //#region CONSTRUCTOR
    constructor(props) {
        super(props);

        this.state = {
            displayType: this.props.displayType,
            listTitle: this.props.list.title
        };
    }
    //#endregion

    //#region RENDER
    public render(): React.ReactElement<IAboutUsAppProps> {
        this.list = this.props.list;
        this.displayType = this.props.displayType;

        return (
            <div className={styles.aboutUsApp}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <span className={styles.title}>List: { this.list.title }</span>
                            <p className={styles.subTitle}>Display as: { this.displayType }</p>
                        </div>
                    </div>
                </div>
                <form>
                    <FormControls field={ this.props.list.fields[0] } displayType='display' />
                </form>
            </div>
        );
    }
    //#endregion
}
