import * as React from 'react';
import styles from './SearchFilter.module.scss';
import { ISearchFilterProps } from './ISearchFilterProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { RxJsEventEmitter } from '../../../libraries/rxJsEventEmitter/RxJsEventEmitter';
import { EventData } from '../../../libraries/rxJsEventEmitter/EventData';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export interface ISearchFilterState {
  eventNumber: number;
}


export default class SearchFilter extends React.Component<ISearchFilterProps, ISearchFilterState> {
  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();

  constructor(props: ISearchFilterProps) {
    super(props);

    this.state = { eventNumber: 0 };
  }


  public render(): React.ReactElement<ISearchFilterProps> {
    return (
      <div className={styles.searchFilter}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {/* <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a> */}
              <h2>ReactiveX Event Emitter</h2>
              <h2>Event Message: {this.state.eventNumber}</h2>
              <PrimaryButton onClick={this.broadcastData.bind(this)} id="BroadcastButton">
                Broadcast message
              </PrimaryButton>

            </div>
          </div>
        </div>
      </div>
    );
  }

  protected broadcastData(): void {

    let eventNumber: number = this.state.eventNumber + 1;

    this.setState((previousState: ISearchFilterState, props: ISearchFilterProps): ISearchFilterState => {
      previousState.eventNumber = eventNumber;
      return previousState;
    });

    this._eventEmitter.emit("myCustomEvent:start", { currentNumber: eventNumber } as EventData);
  }

}
