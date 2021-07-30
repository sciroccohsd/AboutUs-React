import * as React from 'react';
import styles from './AboutUsApp.module.scss';
import { IAboutUsAppProps } from './IAboutUsAppProps';
import { escape } from '@microsoft/sp-lodash-subset';

import DataFactory from './DataFactory';
import CustomDialog from './CustomDialog';


export default class AboutUsApp extends React.Component<IAboutUsAppProps, {}> {
    constructor(props) {
        super(props);

        this.state = {
            displayType: this.props.displayType,
            listTitle: this.props.list.title
        };
    }

    public render(): React.ReactElement<IAboutUsAppProps> {
        return (
            <div className={styles.aboutUsApp}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <span className={styles.title}>List: { this.props.list.title }</span>
                            <p className={styles.subTitle}>Display as: { this.props.displayType }</p>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
