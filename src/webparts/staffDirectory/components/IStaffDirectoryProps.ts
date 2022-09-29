import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldCollectionDataProps } from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";

export interface IStaffDirectoryProps {
  title: string;
  pageSize: number;
  departments: IPropertyFieldCollectionDataProps[];
  showDepartmentFilter: boolean;
  isDarkTheme: boolean;
  context: WebPartContext;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
