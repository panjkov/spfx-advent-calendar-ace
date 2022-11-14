import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'AdventCalendarAdaptiveCardExtensionStrings';
import { IAdventCalendarAdaptiveCardExtensionProps, IAdventCalendarAdaptiveCardExtensionState } from '../AdventCalendarAdaptiveCardExtension';

export class CardView extends BaseImageCardView<IAdventCalendarAdaptiveCardExtensionProps, IAdventCalendarAdaptiveCardExtensionState> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return null;

  }

  public get data(): IImageCardParameters {
    const date = new Date();
    const currentDay = date.getDate();
    const currentMonth = date.getMonth();

      if(currentMonth === 11 && currentDay < 25) {
      return {
        primaryText: this.properties.doorPrefix.concat(" ",this.state.calendarCard.sequence.toString()),
        imageUrl: this.state.calendarCard.pictureUrl ?? require('../assets/MicrosoftLogo.png'),
        title: this.properties.openHereText //this.state.calendarCard.title
      }; 
    }

      else if (currentMonth === 11 && currentDay > 24) {
        return {
          primaryText: strings.CalendarCompleteText,
          imageUrl: "",
          title: ""
        };
    } else {
      return {
        primaryText: strings.CalendarInactiveText,
        imageUrl: "", 
        title: ""
      };
      
    }

  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: this.state.calendarCard.resourceUrl ?? '${this.context.pageContext.absoluteUrl}'
      }
    };
  }
}