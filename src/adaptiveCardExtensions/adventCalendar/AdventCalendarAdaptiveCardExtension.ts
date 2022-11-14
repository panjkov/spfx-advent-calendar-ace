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
  listTitle: string;
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

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());


    this.state = {
      calendarCard: undefined
     };
     
      return this._fetchCurrentCard();

  }

  private _fetchCurrentCard(): Promise<void> {

    const date = new Date();
    const currentDay = date.getDate();


    return this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.properties.listTitle}')/items?$filter=Sequence eq ${currentDay}&$select=Title,Link,Picture,Sequence`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'accept': 'application/json;odata.metadata=none'
      }
    })
      .then(response => response.json())
      .then(calendarCards => {
        const calendarCard = calendarCards?.value.pop();
        let pict = null;
        let title = "";
        let pictureUrl = "";
        let link = null;
        let sequence = currentDay;
        if (calendarCard !== undefined && calendarCard !== null) {
          if(calendarCard.Title !== null) {
            title = calendarCard.Title;
          }

          if(calendarCard.Sequence !== null) {
            sequence = calendarCard.Sequence;
          }
        
        if (calendarCard.Picture !== null) {
          pict = JSON.parse(calendarCard.Picture);
          pictureUrl = pict.serverRelativeUrl;
        }
        else {
          pictureUrl = this.context.pageContext.web.logoUrl;
        }

        if (calendarCard.Link !== null) {
          link = calendarCard.Link.Url;
        }
        else {
          link = this.context.pageContext.web.absoluteUrl;
        }
      }
      else {
        pictureUrl = this.context.pageContext.web.logoUrl;
        link = this.context.pageContext.web.absoluteUrl;
      }

        this.setState({
          calendarCard: {
            title: title || "",
            pictureUrl: pictureUrl ?? "",
            resourceUrl: link,
            sequence: sequence || currentDay
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
