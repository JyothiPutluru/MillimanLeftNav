import ISearchResultSourceData from "../../../models/ISearchResultSourceData";
import IRefinerSourceData from "../../../models/IRefinerSourceData";

export interface IDocumentItem {
    Id: string;
    Url: string;
    Title: string;
    FileName: string;
    PreviewImage: string;
    Description: string;
    Path: string;
    ModifiedBY: string;
    ModifiedDateTime: string;
    Type: string;
}

export interface ITermstoreData {
    Title: string;
    TermDetails: ITerm;
    Terms: ITerm[];
}

export interface ITerm {
    parentId: any;
    Id:string;
    Name:string;
    Owner:string;
    PathOfTerm:string;
    Description:string;
}

export interface ISearchProperties{
    SearchResultsWebPart:string;
    RefinersWebPart:string;
    PaginationWebPart:string;
    SearchBoxWebPart:string;
    SearchVerticalsWebPart:string;
}
export interface IFilterData{
    SearchText:string;
    PathOfTerm:string;
    Refiners:IRefinementFilter[];
    RequiredProp:string;
}
export interface IRefinementValue {
  RefinementCount: number;
  RefinementName: string;
  RefinementToken: string;
  RefinementValue: string;
}

export interface IRefinementFilter {
  FilterName: string;
  Values: IRefinementValue[];
  Operator: RefinementOperator;
}

export enum RefinementOperator {
  OR = 'or',
  AND = 'and'
}

export interface ISearchRecievedProperties{
    Refiners:IRefinementFilter[]|string;
    SearchBoxWebPart:string;
    QueryKeywords:string;
    TermData:ITerm;
    RefinerSourceData:IRefinerSourceData;
}