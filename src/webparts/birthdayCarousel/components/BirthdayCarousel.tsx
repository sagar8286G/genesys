import * as React from 'react';
import styles from './BirthdayCarousel.module.scss';
import customcss from './custom.module.scss';
import { IBirthdayCarouselProps } from './IBirthdayCarouselProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation, CarouselIndicatorShape, ICarouselImageProps } from "@pnp/spfx-controls-react/lib/Carousel";
import { ICssInput, ImageFit, styled } from 'office-ui-fabric-react';
import spservices, { IAllItems } from '../../../Services/spService';
import { stringIsNullOrEmpty } from '@pnp/common';

export interface IcarouselElement {
  imageSrc: string;
  title: string;
  description: any;
  // url: string;
  showDetailsOnHover: boolean;
  imageFit: any;
}

export interface IBirthdayCarouselStates {
  carouselElement: IcarouselElement[];
}

export default class BirthdayCarousel extends React.Component<IBirthdayCarouselProps, IBirthdayCarouselStates> {
  private ServiceInstance: spservices = null;
  public constructor(props: IBirthdayCarouselProps) {
    super(props);
    this.ServiceInstance = new spservices(this.props.context);

    this.state = {
      carouselElement: []
    };
  }

  public getISODateString(searchDate: Date, searchTime: string) {
    let returnString = '';
    try {
      if (searchDate && !stringIsNullOrEmpty(searchTime)) {
        let tempTime = searchTime.split(':');
        returnString = new Date(searchDate.setHours(parseInt(tempTime[0]), parseInt(tempTime[1]), 0, 0)).toISOString();
      }
    }
    catch (error) {
      throw error;
    }
    return returnString;
  }

  public async componentDidMount() {
    let temp: IcarouselElement[] = [];
    let AllItemQuery: IAllItems = {
      listName: this.props.listName,
      filterQuery: `(Birthday ge '${this.getISODateString(new Date(), "00:00")}' and Birthday le '${this.getISODateString(new Date(), "23:59")}')
      or (Anniversary ge '${this.getISODateString(new Date(), "00:00")}' and Anniversary le '${this.getISODateString(new Date(), "23:59")}')`
    };
    let occation = await this.ServiceInstance.getAllListItems(AllItemQuery);

    if (occation.length > 0) {
      occation.map(item => {
        let isBirthday = false;
        let isAnniversary = false;
        let birthDate = item.Birthday ? new Date(item.Birthday).getDate() : null;
        let birthMonth = item.Birthday ? new Date(item.Birthday).getMonth() + 1 : null;
        let anniversaryDate = item.Anniversary ? new Date(item.Anniversary).getDate() : null;
        let anniversaryMonth = item.Anniversary ? new Date(item.Anniversary).getMonth() + 1 : null;
        isBirthday = (new Date().getDate() === birthDate && new Date().getMonth() + 1 === birthMonth);
        isAnniversary = (new Date().getDate() === anniversaryDate && new Date().getMonth() + 1 === anniversaryMonth);
        const descrptionHTML: JSX.Element =
          <div>
            <p className={customcss.datePosition}>{new Date().toLocaleDateString()}</p>
            <p>{isBirthday ? item.BirthdayMessage : isAnniversary ? item.AnniversaryMessage : ''}</p>
          </div>

        temp.push({
          imageSrc: require(isBirthday ? '../assets/Birthday.png' : isAnniversary ? '../assets/Anniversary.gif' : '../assets/Birthday.png'),
          title: `${item.Title}`,
          // url: `${new Date().toLocaleDateString()}`,
          // description: isBirthday ? 'Happy Birthday' : isAnniversary ? 'Happy Anniversary' : '',
          description: descrptionHTML,
          showDetailsOnHover: false,
          imageFit: ImageFit.cover
        })
      });
    } else {
      temp.push({
        imageSrc: require('../assets/EmptyImage.jpg'),
        title: `No Event`,
        description: '',
        showDetailsOnHover: false,
        imageFit: ImageFit.cover
      })
    }

    this.setState({ carouselElement: temp });
  }

  public async componentDidUpdate(pp: IBirthdayCarouselProps, ps: IBirthdayCarouselStates) {
    if (ps.carouselElement !== this.state.carouselElement) {

    }
  }

  public render(): React.ReactElement<IBirthdayCarouselProps> {
    return (
      <div>
        <Carousel
          buttonsLocation={CarouselButtonsLocation.top}
          buttonsDisplay={CarouselButtonsDisplay.block}

          contentContainerStyles={customcss.customContainer}
          // containerButtonsStyles={styles.carouselButtonsContainer}
          indicatorShape={CarouselIndicatorShape.circle}
          interval={3000}
          slide={true}
          isInfinite={true}
          element={this.state.carouselElement}

          onMoveNextClicked={(index: number) => { }}
          onMovePrevClicked={(index: number) => { }}
        />
      </div >
    );
  }
}
