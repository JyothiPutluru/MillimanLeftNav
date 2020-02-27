import { WebPartContext } from "@microsoft/sp-webpart-base";
import ISearchResultSourceData from "../../../models/ISearchResultSourceData";
import { IRefinementFilter } from "../../../models/ISearchResult";
import { ISearchRecievedProperties } from "./pgp.models";
import IRefinerSourceData from "../../../models/IRefinerSourceData";

export interface IPgpLeftNavigationProps {
  context:WebPartContext;
  description: string;
  taxonomyTermstoreId: string;
  taxonomyTermGroupId:string;
  taxonomyTermsetId: string;
  taxonomyTermId: string;
  mappedField: string;
  showTitle: boolean;
  title: string;
  titleLinkUrl: string;
  isTreeView: boolean;
  searchQuery:string;
  refinerSourceData:IRefinerSourceData;
  queryKeywords:string;
  defaultQueryKeywords:string;
  transmitSearchProperties:(changedProperties:ISearchRecievedProperties )=>void;
}
