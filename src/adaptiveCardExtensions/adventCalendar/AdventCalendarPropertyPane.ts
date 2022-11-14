import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'AdventCalendarAdaptiveCardExtensionStrings';

export class AdventCalendarPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleLabel
                }),
                PropertyPaneTextField('doorPrefix', {
                  label: strings.DoorPrefixLabel
                }),
                PropertyPaneTextField('openHereText', {
                  label: strings.OpenHereTextLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
