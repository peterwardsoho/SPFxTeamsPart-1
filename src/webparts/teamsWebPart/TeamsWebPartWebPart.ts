import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TeamsWebPartWebPart.module.scss';
import * as strings from 'TeamsWebPartWebPartStrings';
import * as microsoftTeams from '@microsoft/teams-js';


export interface ITeamsWebPartWebPartProps {
  description: string;
}

export default class TeamsWebPartWebPart extends BaseClientSideWebPart<ITeamsWebPartWebPartProps> {

  private teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this.teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    let title: string = '';
    let siteTabTitle: string = '';
    if (this.teamsContext) {
      title = "Welcome to MS Teams!";
      siteTabTitle = "Team: " + this.teamsContext.teamName;
    }
    else {
      title = "Welcome to SharePoint!";
      siteTabTitle = "SharePoint site: " + this.context.pageContext.web.title;
    }
    this.domElement.innerHTML = `
      <div class="${ styles.teamsWebPart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
          <div class="${ styles.column}">  
          <span class="${ styles.title}">${title}</span>  
          <p class="${ styles.description}">${siteTabTitle}</p>  
          <p class="${ styles.description}">Description property value - ${escape(this.properties.description)}</p>  
          <a href="https://aka.ms/spfx" class="${ styles.button}">  
            <span class="${ styles.label}">Learn more</span>  
          </a>  
        </div>  
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
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
