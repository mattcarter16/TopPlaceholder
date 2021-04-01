import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { IHeaderProps, Header } from './components/Header';

import * as strings from 'TopPlaceholderApplicationCustomizerStrings';

const LOG_SOURCE: string = 'TopPlaceholderApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITopPlaceholderApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class TopPlaceholderApplicationCustomizer
  extends BaseApplicationCustomizer<ITopPlaceholderApplicationCustomizerProperties> {

    private static headerPlaceholder: PlaceholderContent;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.application.navigatedEvent.add(this, () => {
      const cssUrl: string = "https://aptitude4dev.sharepoint.com/Shared%20Documents/Header.css";
    const url = window.location.href;

    if(cssUrl) {
      console.log(window.location.href);
      const head: any =document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }
    else {
      const links=document.getElementsByTagName("link");
      console.log(links);
      for(let i=0; i < links.length; i++) {
        if(links[i].href.indexOf(cssUrl)>-1){
          links[i].remove();
        }
      }
    }

      this.loadReactComponent();
    });
  
    this._render();

    return Promise.resolve();
  }

  private _render() {
    if (this.context.placeholderProvider.placeholderNames.indexOf(PlaceholderName.Top) !== -1) {
      if (!TopPlaceholderApplicationCustomizer.headerPlaceholder || !TopPlaceholderApplicationCustomizer.headerPlaceholder.domElement) {
        TopPlaceholderApplicationCustomizer.headerPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
          onDispose: this.onDispose
        });
      }

      this.loadReactComponent();
    }
    else {
      console.log(`The following placeholder names are available`, this.context.placeholderProvider.placeholderNames);
    }
  }

  private loadReactComponent() {
    if (TopPlaceholderApplicationCustomizer.headerPlaceholder && TopPlaceholderApplicationCustomizer.headerPlaceholder.domElement) {
      const element: React.ReactElement<IHeaderProps> = React.createElement(Header, {
        context: this.context
      });

      ReactDom.render(element, TopPlaceholderApplicationCustomizer.headerPlaceholder.domElement);
    }
    else {
      console.log('DOM element of the header is undefined. Start to re-render.');
      this._render();
    }
  }

  private _onDispose(): void {
    console.log('[PreallocateSpaceApplicationCustomizer._onDispose] Disposed custom top placeholder.');
  }
}
