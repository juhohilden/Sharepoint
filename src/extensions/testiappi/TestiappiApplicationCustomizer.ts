import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent, 
  PlaceholderName, 
} from '@microsoft/sp-application-base';

import * as strings from 'TestiappiApplicationCustomizerStrings';
import styles from './TestiappiApplicationCustomizer.module.scss';
import {escape} from '@microsoft/sp-lodash-subset';
//import { PlaceholderContent } from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'TestiappiApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITestiappiApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  Top : string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Content : string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class TestiappiApplicationCustomizer
  extends BaseApplicationCustomizer<ITestiappiApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
    console.log(
      "Available placeholders: ",
      this.context.placeholderProvider.placeholderNames
        .map(name => PlaceholderName[name])
        .join(", ")
    );
    
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }
  
      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = "(Top property was not defined.)";
        }
  
        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.top}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape('Testi Header')}
            </div>
          </div>`;
        }
      }
    }
  }

  private _onDispose(): void {
    console.log('[TestiappidApplicationCustomizer._onDispose] Disposed custom top.');
  }
}
