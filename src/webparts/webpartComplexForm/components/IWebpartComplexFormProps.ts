import { SPHttpClient } from '@microsoft/sp-http';

export interface IWebpartComplexFormProps {
  spHttpClient: SPHttpClient;
  description: string;
  ProjectName: string;
  ProjectsArray: Array<string>[];
  siteurl: string;
  Building: string;
  Floor: string;
  GridLine: string;
  Subject: string;
  SubContractor: string;
  CreatedDate: string;
  RequiredDate: string;
  Disciplined: string;
  Description: string;
  Requirement: string;
  Comments: string;
  ItemGuid:string;
  loading: boolean;
  UploadedFilesArray: Array<string>[];
  CurrentUser: string;
  UserGroup: string;
  IRFINumber: string;
  IRFISeriesId: string;
  IRFIReference: string;
}
