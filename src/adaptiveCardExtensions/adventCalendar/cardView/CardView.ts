import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'AdventCalendarAdaptiveCardExtensionStrings';
import { IAdventCalendarAdaptiveCardExtensionProps, IAdventCalendarAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../AdventCalendarAdaptiveCardExtension';

export class CardView extends BaseImageCardView<IAdventCalendarAdaptiveCardExtensionProps, IAdventCalendarAdaptiveCardExtensionState> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IImageCardParameters {
    return {
      primaryText: this.properties.doorPrefix.concat(" ",this.state.calendarCard.sequence.toString()),
      imageUrl: this.state.calendarCard.pictureUrl ?? require('../assets/MicrosoftLogo.png'),
      title: this.properties.openHereText //this.state.calendarCard.title
    };
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
// require('../assets/MicrosoftLogo.png')