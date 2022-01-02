import * as React from 'react';
import styles from './Spfx3Wp1.module.scss';
import { ISpfx3Wp1Props } from './ISpfx3Wp1Props';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Spfx3Wp1 extends React.Component<ISpfx3Wp1Props, {}> {
  public render(): React.ReactElement<ISpfx3Wp1Props> {
    return (
      <div className={ styles.spfx3Wp1 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
