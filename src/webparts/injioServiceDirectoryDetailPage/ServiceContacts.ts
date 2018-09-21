import { EmailAddress } from "@microsoft/microsoft-graph-types";
export interface ServiceContacts {
    value: ServiceContact[];
  }


export interface ServiceContact{
    FirstName:string;
    LastName:string;
    JobTitle:string;
    BusinessPhone:string;
    EMail:string;
}