import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDynamicDropdownProps {
  description: string;
  context: WebPartContext;
  siteUrl:string;
  singleValueOptions:any;
  multiValueOptions:any;
}
