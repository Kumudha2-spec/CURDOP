import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CurdOperationsSpfxWebPartStrings';
import CurdOperationsSpfx from './components/CurdOperationsSpfx';
import { ICurdOperationsSpfxProps } from './components/ICurdOperationsSpfxProps';

export interface ICurdOperationsSpfxWebPartProps {
  description: string;
  listName: string;
  defaultRating: number;
  showRating: boolean;
  allowAnonymous: boolean;
}

export default class CurdOperationsSpfxWebPart
  extends BaseClientSideWebPart<ICurdOperationsSpfxWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICurdOperationsSpfxProps> =
      React.createElement(
        CurdOperationsSpfx,
        {
          description: this.properties.description || 'Employee Feedback',
          listName: this.properties.listName || 'EmployeeFeedback',
          defaultRating: this.properties.defaultRating || 3,
          showRating: this.properties.showRating !== false,
          allowAnonymous: this.properties.allowAnonymous || false,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context
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
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';

          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;

            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;

            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;

            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(
    currentTheme: IReadonlyTheme | undefined
  ): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        '--bodyText',
        semanticColors.bodyText || null
      );

      this.domElement.style.setProperty(
        '--link',
        semanticColors.link || null
      );

      this.domElement.style.setProperty(
        '--linkHovered',
        semanticColors.linkHovered || null
      );
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
            description: 'Employee Feedback Form Configuration'
          },

          groups: [
            {
              groupName: 'Form Settings',

              groupFields: [
                // Form Title / Heading
                PropertyPaneTextField('description', {
                  label: 'Form Heading'
                }),

                // Target SharePoint List
                PropertyPaneTextField('listName', {
                  label: 'Target SharePoint List Name'
                }),

                // Default Star/Number Rating slider
                PropertyPaneSlider('defaultRating', {
                  label: 'Default Rating (1 to 5)',
                  min: 1,
                  max: 5,
                  step: 1
                }),

                // Toggle visibility of the Rating section
                PropertyPaneCheckbox('showRating', {
                  text: 'Enable Rating Scale'
                }),

                // Allow anonymous submissions (hides name & email)
                PropertyPaneCheckbox('allowAnonymous', {
                  text: 'Allow Anonymous Submissions'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}