import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { AdventCalendarPropertyPane } from './AdventCalendarPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IAdventCalendarAdaptiveCardExtensionProps {
  title: string;
  doorPrefix: string;
  openHereText: string;
}

export interface IAdventCalendarAdaptiveCardExtensionState {
  calendarCard: ICalendarCard | undefined;
}

export interface ICalendarCard {
  title: string;
  pictureUrl: string;
  resourceUrl: string;
  sequence: number;
}

const CARD_VIEW_REGISTRY_ID: string = 'AdventCalendar_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'AdventCalendar_QUICK_VIEW';

export default class AdventCalendarAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAdventCalendarAdaptiveCardExtensionProps,
  IAdventCalendarAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AdventCalendarPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      calendarCard: undefined
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchCurrentCard();
  }

  private _fetchCurrentCard(): Promise<void> {
    const date = new Date();
    const currentDay = date.getDate();


    return this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('AdventCalendar')/items?$filter=Sequence eq ${currentDay}&$select=Title,Link,Picture,Sequence`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'accept': 'application/json;odata.metadata=none'
      }
    })
      .then(response => response.json())
      .then(calendarCards => {
        const calendarCard = calendarCards.value.pop();
        let pict = null;
        let pictureUrl = "";
        if (calendarCard.Picture !== null) {
          pict = JSON.parse(calendarCard.Picture);
          pictureUrl = pict.serverRelativeUrl;
        }
        else {
          pictureUrl = this.context.pageContext.web.logoUrl;
        }

        let link = null;
        if (calendarCard.Link !== null) {
          link = calendarCard.Link.Url;
        }
        else {
          link = this.context.pageContext.web.absoluteUrl;
        }
        
        // var s1 = new String(calendarCard.Picture.serverUrl);
        // var s2 = new String(calendarCard.Picture.serverRelativeUrl);
        // var s3 = s1.concat(s2.toString());

        this.setState({
          calendarCard: {
            title: calendarCard.Title,
            pictureUrl: pictureUrl ?? "",
            resourceUrl: link,
            sequence: calendarCard.Sequence
          }
        })
      })
      .catch(error => console.error(error));
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'AdventCalendar-property-pane'*/
      './AdventCalendarPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.AdventCalendarPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  protected get iconProperty(): string {
    return 'snowflake';
  }
}
