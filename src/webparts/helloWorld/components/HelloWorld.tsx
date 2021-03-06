import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  renderItemTemplate() {
    return (
      this.props.itemTemplate
    )
  }
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
        <div dangerouslySetInnerHTML={{__html: this.props.itemTemplate}} />

        </div>
      </div>

    );
  }
}
