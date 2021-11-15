import * as React from 'react';
import styles from './ContentAccordion.module.scss';
import { IContentAccordionProps } from './IContentAccordionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices, { IAllItems } from '../../../Services/spService';
import './reactAccordion.css';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';

export interface IContentAccordionStates {
  items: Array<any>;
  allowMultipleExpanded: boolean;
  allowZeroExpanded: boolean;

}

export default class ContentAccordion extends React.Component<IContentAccordionProps, IContentAccordionStates> {
  private ServiceInstance: spservices = null;
  public constructor(props: IContentAccordionProps) {
    super(props);
    this.ServiceInstance = new spservices(this.props.context);

    this.state = {
      items: new Array<any>(),
      allowMultipleExpanded: this.props.allowMultipleExpanded,
      allowZeroExpanded: this.props.allowZeroExpanded
    };
  }

  public componentDidUpdate(prevProps: IContentAccordionProps): void {
    // if(prevProps.listName !== this.props.listName) {
    //   this.getListItems();
    // }

    // if(prevProps.allowMultipleExpanded !== this.props.allowMultipleExpanded || prevProps.allowZeroExpanded !== this.props.allowZeroExpanded) {
    //   this.setState({
    //     allowMultipleExpanded: this.props.allowMultipleExpanded,
    //     allowZeroExpanded: this.props.allowZeroExpanded
    //   });
    // }
  }

  public async componentDidMount() {
    let AllItemQuery: IAllItems = {
      listName: this.props.listName
    };
    let accordionItems = await this.ServiceInstance.getAllListItems(AllItemQuery);
    this.setState({ items: accordionItems });
  }

  public render(): React.ReactElement<IContentAccordionProps> {
    const { allowMultipleExpanded, allowZeroExpanded } = this.state;

    const accordionHTML: JSX.Element[] = this.state.items.map(item => {
      let dummyElement = document.createElement('div');
      dummyElement.innerHTML = item.RichContent;
      let outputText = dummyElement.innerText;
      return (
        <AccordionItem>
          <AccordionItemHeading>
            <AccordionItemButton>
              {item.Title}
            </AccordionItemButton>
          </AccordionItemHeading>
          <AccordionItemPanel>
            <p dangerouslySetInnerHTML={{ __html: `${outputText}` }} />
          </AccordionItemPanel>
        </AccordionItem>
      )
    });

    return (
      <div className={styles.contentAccordion}>
        <div>
          <h3>{this.props.accordionTitle}</h3>
          {
            accordionHTML.length > 0 &&
            <Accordion allowZeroExpanded={allowZeroExpanded} allowMultipleExpanded={allowMultipleExpanded}>
              {
                accordionHTML
              }
            </Accordion>
          }
        </div>
      </div>
    );
  }
}
