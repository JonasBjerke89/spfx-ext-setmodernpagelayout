import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderName,
  PlaceholderContent
} from '@microsoft/sp-application-base';

import * as strings from 'SetModernPageLayoutApplicationCustomizerStrings';
import { ISetModernPageLayoutProps } from './components/ISetModernPageLayoutProps';
import SetModernPageLayoutComponent from './components/SetModernPageLayoutComponent';
import { sp } from "@pnp/sp";

const LOG_SOURCE: string = 'SetModernPageLayoutApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISetModernPageLayoutApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SetModernPageLayoutApplicationCustomizer
  extends BaseApplicationCustomizer<ISetModernPageLayoutApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    sp.setup({
      spfxContext: this.context
    });

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }
      
      const element: React.ReactElement<ISetModernPageLayoutProps> = React.createElement(
        SetModernPageLayoutComponent,
        {
          context: this.context
        }
      );
  
      ReactDom.render(element, this._topPlaceholder.domElement);
    }

    return Promise.resolve();
  }
}
