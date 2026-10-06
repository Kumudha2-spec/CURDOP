import * as React from 'react';
import styles from './CurdOperationsSpfx.module.scss';
import type { ICurdOperationsSpfxProps } from './ICurdOperationsSpfxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import welcomeDark from '../assets/welcome-dark.png';
import welcomeLight from '../assets/welcome-light.png';

export default class CurdOperationsSpfx extends React.Component<ICurdOperationsSpfxProps> {
  public render(): React.ReactElement<ICurdOperationsSpfxProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.curdOperationsSpfx}`}>
        <div className={styles.welcome}>
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
         
        </div>
      </section>
    );
  }
}
