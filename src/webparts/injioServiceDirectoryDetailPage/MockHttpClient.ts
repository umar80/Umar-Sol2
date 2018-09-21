import {ServiceDirectory} from './ServiceDirectoryList';
import {ServiceComment} from './ServiceComments';
import {ServiceContact} from './ServiceContacts';
import {AdditionalAttribute} from './ServiceDirectoryAdditionalAttributes';
import { DateTimeColumn } from "@microsoft/microsoft-graph-types";

export default class MockHttpClient{

    private static _serviceDirectoryItem: ServiceDirectory[] =[     {
        Title: 'Mock List',
        ID: 1 ,
        AverageRating:5,
        Contact:"Umar Riaz",
        Description:"This is a record for one of the Service Provider. This is a record for one of the Service Provider. This is a record for one of the Service Provider. This is a record for one of the Service Provider. This is a record for one of the Service Provider.",
        LocationMap:'https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3311.926234334049!2d151.20472941528772!3d-33.891553580649926!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x6b12b1defbc4a479%3A0xfa389e4e869315ec!2s59+Great+Buckingham+St%2C+Redfern+NSW+2016!5e0!3m2!1sen!2sau!4v1536298216548',
        Logo:"https://webvine.sharepoint.com/sites/MIqbalTest/PublishingImages/Lists/Service%20Directory/AllItems/asus.jpg",
        Phone:"111111111111",
        ServiceType:"Engineering",
        Website:"http://google.com",
        Country:"Australia",
        State:"NSW",
        Region:"Australia",
        Centre:"REDFERN",
        Email:"test@test.com",
        ABN:"1234566788555",
        Address:"123-456 Street, street1",
        Status:"Active"
   }];


    
    private static _serviceDirectoryAdditionalAttributeItem: AdditionalAttribute[] =[ {
        Service:"Mock List",
        StartDate:"06/09/18",
        EndDate:"06/09/19",
        NoticePeriodDate: "01/09/18",
        PrimaryContact:"Umar Riaz"
   }];


   private static _serviceDirectoryComments: ServiceComment[]= [{
        Service:"Mock List",
        Comments:"Test comments",
        Created: new Date(2018,12,22,12,0,1) ,
        CreatedBy:"Umar Riaz"

   }];

   private static _serviceDirectoryContacts: ServiceContact[] =[ {
        FirstName:"Umar",
        LastName:"Riaz",
        JobTitle:"Sr Sharepoint Consultant",
        BusinessPhone:"111111111",
        EMail:"umar@webvine.com.au" 

   }];

    public static getServiceDirectoryItem(): Promise<ServiceDirectory[]> {
        return new Promise<ServiceDirectory[]>((resolve) => {
                resolve(MockHttpClient._serviceDirectoryItem);
            });
        }

    public static getAdditionalAttributes(): Promise<AdditionalAttribute[]> {
            return new Promise<AdditionalAttribute[]>((resolve) => {
                    resolve(MockHttpClient._serviceDirectoryAdditionalAttributeItem);
                });
              }


    public static getServiceContacts(): Promise<ServiceContact[]> {
            return new Promise<ServiceContact[]>((resolve) => {
                    resolve(MockHttpClient._serviceDirectoryContacts);
                });
              }


     public static getServiceComments(): Promise<ServiceComment[]> {
            return new Promise<ServiceComment[]>((resolve) => {
                    resolve(MockHttpClient._serviceDirectoryComments);
                });
              }
}