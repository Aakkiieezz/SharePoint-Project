import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'Spfx3Wp1WebPartStrings';
import Spfx3Wp1 from './components/Spfx3Wp1';
import Status from './components/Status';
import { ISpfx3Wp1Props } from './components/ISpfx3Wp1Props';
export interface ISpfx3Wp1WebPartProps { description: string; }

export default class Spfx3Wp1WebPart extends BaseClientSideWebPart<ISpfx3Wp1WebPartProps> {

  public render(): void
  {
    const element: React.ReactElement<ISpfx3Wp1Props> = React.createElement(
      Status,
      {
        description: this.properties.description
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void
  {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version
  {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration
  {
    return{
      pages:
      [{
          header:{description: strings.PropertyPaneDescription},
          groups:
          [{
              groupName: strings.BasicGroupName,
              groupFields: [PropertyPaneTextField("description",{label: strings.DescriptionFieldLabel})]
          }]
      }]
    };
  }
}