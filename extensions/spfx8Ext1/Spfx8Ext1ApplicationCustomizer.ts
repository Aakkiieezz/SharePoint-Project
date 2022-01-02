import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'Spfx8Ext1ApplicationCustomizerStrings';

const LOG_SOURCE: string = 'Spfx8Ext1ApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfx8Ext1ApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class Spfx8Ext1ApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfx8Ext1ApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    //Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    let bottomPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom); 
    if(bottomPlaceholder)
      bottomPlaceholder.domElement.innerHTML = `
      <div style="background-color:rgba(24,37,52,255); text-align:right;">
        <marquee direction="left">©2021 shop.adidas.co.in | Powered By : Adi Sports (India) Pvt. Ltd.</marquee>
      </div>`;

    return Promise.resolve();
  }
}
