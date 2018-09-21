import { DateTimeColumn } from "@microsoft/microsoft-graph-types";
export interface AdditionalAttributes
{
    value: AdditionalAttribute[];
}


export interface AdditionalAttribute{
    Service:string;
    StartDate:string;
    EndDate:string;
    NoticePeriodDate:string;
    PrimaryContact:string;

}


