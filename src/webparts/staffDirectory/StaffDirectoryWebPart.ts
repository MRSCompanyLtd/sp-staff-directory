import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { PropertyFieldCollectionData, CustomCollectionFieldType, IPropertyFieldCollectionDataProps } from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'StaffDirectoryWebPartStrings';
import StaffDirectory from './components/StaffDirectory';
import { IStaffDirectoryProps } from './components/IStaffDirectoryProps';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls';

export interface IStaffDirectoryWebPartProps {
  title: string;
  pageSize: number;
  departments: IPropertyFieldCollectionDataProps[];
  showDepartmentFilter: boolean;
}

export default class StaffDirectoryWebPart extends BaseClientSideWebPart<IStaffDirectoryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IStaffDirectoryProps> = React.createElement(
      StaffDirectory,
      {
        title: this.properties.title,
        pageSize: this.properties.pageSize,
        departments: this.properties.departments,
        showDepartmentFilter: this.properties.showDepartmentFilter,
        isDarkTheme: this._isDarkTheme,
        context: this.context,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this.properties.title = 'Staff Directory';
    this.properties.pageSize = 10;
    this.properties.showDepartmentFilter = false;
    this.properties.departments = [];

    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneSlider('pageSize', {
                  label: 'Results per page',
                  showValue: true,
                  max: 20,
                  min: 4,
                  step: 2,
                  value: this.properties.pageSize
                }),
                PropertyFieldCollectionData('departments', {
                  key: 'departments',
                  label: 'Department List',
                  panelHeader: 'Department List',
                  manageBtnLabel: 'Manage List',
                  value: this.properties.departments,
                  fields: [
                    {
                      id: 'departmentKey',
                      title: 'Department key from AD',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'departmentName',
                      title: 'Display name for department',
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ]
                }),
                PropertyFieldToggleWithCallout('showDepartmentFilter', {
                  key: 'showDepartmentFilter',
                  label: 'Show Department Filter',
                  onText: 'Yes',
                  offText: 'No',
                  onAriaLabel: 'Department filter on',
                  offAriaLabel: 'Department filter off',
                  checked: this.properties.showDepartmentFilter
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
