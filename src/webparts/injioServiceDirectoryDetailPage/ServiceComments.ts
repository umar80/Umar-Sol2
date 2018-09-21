import { DateTimeColumn } from "@microsoft/microsoft-graph-types";
export interface ServiceComments
{
   value: ServiceComment[];

}


export interface ServiceComment{

    Service:string;
    Comments:string;
    Created:Date;
    CreatedBy:string;
}