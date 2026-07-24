import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PrtsWebPartStrings';
import Mdr from './components/Prts';
import { IPrtsProps } from './components/IPrtsProps';
import IEmployeeProfileops from './service/BAL/SPCRUD/EmployeeProfile';
import SPCRUDOPS from "./service/DAL/spcrudops";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMdrWebPartProps {
  description: string;
}

export default class MdrWebPart extends BaseClientSideWebPart<IMdrWebPartProps> {


  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {

    const IASACL = async () => {
      const spCrudOps = await SPCRUDOPS();
      const IASdata = await spCrudOps.getData(
        'PRTSACL',
        'Title,UserName/EMail,Role,Status',
        'UserName',
        '',
        { column: 'ID', isAscending: true },
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          currentSPContext: this.context,
          userEmail: this.context.pageContext.user.email
        }
      );

      let AAfiltereddata = IASdata.filter((m) => m.UserName?.EMail === this.context.pageContext.user.email);

      let Maintfiltereddata = IASdata.filter((m) => m.Title === "Maintenance")

      // Default values
      let isAppAdmin: boolean = false;
      let isMaintenance: boolean = false;
      let isSysAdmin: boolean = false;
      let isReader: boolean = false;

      if (AAfiltereddata[0]?.Title === 'AppAdmin') {
        isAppAdmin = true;
      }
      if (AAfiltereddata[0]?.Title === 'SysAdmin') {
        isSysAdmin = true;
      }
      if (Maintfiltereddata[0]?.Status === 'Active') {
        isMaintenance = true;
      }
      if (AAfiltereddata[0]?.Role === 'Reader') {
        isReader = true;
      }

      return { isAppAdmin, isMaintenance, isSysAdmin, isReader };
    };

    // ✅ Get ACL values
    const { isAppAdmin, isMaintenance, isSysAdmin, isReader } = await IASACL();


    const itemdata = await IEmployeeProfileops().getEmployeeProfile(this.context.pageContext.user.email, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      currentSPContext: this.context,
      userEmail: this.context.pageContext.user.email,
      context: this.context
      //EmployeeId:''
    });
    const element: React.ReactElement<IPrtsProps> = React.createElement(
      Mdr,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        currentSPContext: this.context,
        userEmail: this.context.pageContext.user.email,
        EmployeeId: itemdata,
        context: this.context,
        Appadmin: isAppAdmin,
        SysAdmin: isSysAdmin,
        Maintenance: isMaintenance,
        Reader: isReader
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
