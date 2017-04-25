import {
  Version,
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './QuickLinks.module.scss';
import * as strings from 'quickLinksStrings';
import { IQuickLinksWebPartProps } from './IQuickLinksWebPartProps';
import 'jquery';
import pnp from 'sp-pnp-js';

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {

  public render(): void {
    if (Environment.type === EnvironmentType.Local) {
      // this.domElement.innerHTML = this.getLinksHTML();
      this.domElement.innerHTML = `
      <div class="${styles.quickLinks}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Quick Links (Running Locally)</span>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.LinkSet)}</p>
            </div>
          </div>
        </div>
      </div>`;
    }
    else if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this.domElement.innerHTML = `
      <div class="${styles.quickLinks}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Quick Links (Running in SPO)</span>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.LinkSet)} Links</p>
              <div id="qlinks" />
            </div>
          </div>
        </div>
      </div>`;

      $('#qlinks', this.domElement).html(this.getLinkData());
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getLinkData(): string {
    let linkData: string = '';
    let linkFilter: string = `LinkSet eq '${this.properties.LinkSet}' and Visible eq 1`;

    pnp.sp.web.lists.getByTitle('QuickLinks').items.filter(linkFilter).orderBy("Sequence").get().then(r => {

      let row: string = '';

      // start the unordered list
      linkData = "<ul>";

      // loop through the returned items to create the links
      r.array.forEach(item => {
        row += "<li><a href='";
        row += item.URL.Url;
        row += "'";

        // if external, open in new window/tab
        if (item.External) {
          row += " target='_blank'";
        }

        // if there are comments, turn them into a tooltip
        if (item.Comments.Length > 0) {
          row += " title='";
          row += item.Comments;
          row += "'";
        }

        //set the link text
        row += ">";
        row += item.URL.Description;

        // if extnernal, add the indicator icon
        if (item.External) {
          row += " <i class='ms-Icon ms-Icon--OpenInNewWindow' aria-hidden='true'></i>";
        }

        // close the link
        row += "</a></li>";

        // add the link to the list of links
        linkData += row;
        row = '';
      });

      // close the link list
      linkData += "</ul>";
    });

    return linkData;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('LinkSet', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
