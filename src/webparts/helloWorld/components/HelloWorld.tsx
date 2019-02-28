import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';

export default class HelloWorld extends React.Component<IHelloWorldProps, any> {

  constructor(props) {
    super(props);
    this.state = {
      items: []
    };

  }

  private renderItemTemplate() {
    return (
      this.props.itemTemplate
    );
  }
  public componentDidMount() {
    console.log("****************8mounted");
    if (this.props.selectedFields !== undefined) {
      console.log(this.props.selectedFields.toString());
      sp.web.lists.getById(this.props.selectedList).select(this.props.selectedFields.toString()).items.get().then(items => {

        this.setState({ items: items });
        console.log(this.state.items);
      });
    }


  }
  public render(): React.ReactElement<IHelloWorldProps> {

    let items=this.state.items.map(item => {
      return (
        this.props.selectedFields.map(field => {
        console.log('field:'+ item[field]);
        return (
          <div>
            {
              item[field]
            }

          </div>
        );
      }));
    });
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div>
            {
              items
            }

          </div>


        </div>
      </div>

    );
  }
}
