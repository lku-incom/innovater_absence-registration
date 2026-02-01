import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AbsenceRegistration from './components/AbsenceRegistration';
import { IAbsenceRegistrationProps } from './components/IAbsenceRegistrationProps';

export interface IAbsenceRegistrationWebPartProps {
  title: string;
  dataverseUrl: string;
}

export default class AbsenceRegistrationWebPart extends BaseClientSideWebPart<IAbsenceRegistrationWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAbsenceRegistrationProps> =
      React.createElement(AbsenceRegistration, {
        context: this.context,
        title: this.properties.title || 'Fraværsregistrering',
        dataverseUrl: this.properties.dataverseUrl,
      });

    ReactDom.render(element, this.domElement);
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
            description: 'Konfiguration af fraværsregistrering',
          },
          groups: [
            {
              groupName: 'Indstillinger',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Titel',
                  value: 'Fraværsregistrering',
                }),
                PropertyPaneTextField('dataverseUrl', {
                  label: 'Dataverse URL',
                  description:
                    'URL til dit Dataverse miljø (f.eks. https://org.crm4.dynamics.com)',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
