import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDocumentItem, ITermstoreData, ITerm } from '../components/pgp.models';
import { Session } from "@pnp/sp-taxonomy";
import PGPConstants from './pgp.constants';

export interface IPGPSerive {
    GetSearchResults(searchText: string): Promise<IDocumentItem[]>;
    GetTermstoreData(termstoreId: string, termGroupId: string, termsetId: string, termId: string): Promise<ITermstoreResult>;
}

export interface ITermstoreResult {
    GroupedData: ITermstoreData[];
    UnGroupedData:ITerm[];
}

export default class PGPSerive implements IPGPSerive {
    private ctx: WebPartContext;
    constructor(context: WebPartContext) {
        this.ctx = context;
    }

    public async GetSearchResults(searchText: string): Promise<IDocumentItem[]> {
        var documentItems: IDocumentItem[] = [];
        let queryText: string = `querytext=''`;
        let url: string = `${this.ctx.pageContext.web.absoluteUrl}/_api/search/query?`;
        url += queryText;
        url += `&queryTemplate='{searchTerms} SiteId:"` + this.ctx.pageContext.site.id + `" AND IsDocument:1 NOT IsOneNotePage:1 NOT(ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118*) NOT(IsContainer :1) NOT(FileExtension:aspx) NOT(FileExtension:html)'`;
        url += `&SelectProperties='ContentClass,PictureThumbnailURL,Description,ContentTypeId,DefaultEncodingURL,DocId,EditorOWSUSER,FileExtension,GeoLocationSource,LastModifiedTime,ModifiedBy,Path,SPWebUrl,SecondaryFileExtension,LinkingUrl,ServerRedirectedUrl,SiteId,Title,Filename,UniqueId,WebId'`;
        url += `&SortList='LastModifiedTime:descending'`;
        try {
            const resp = await this.ctx.spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: {
                    'odata-version': '3.0',
                    'accept': 'application/json;odata=verbose',
                    'content-type': 'application/json;odata=verbose'
                }
            });
            if (resp.ok) {
                return resp.json().then((data) => {
                    var results = this.getDataFromQueryResults(data.d.query);
                    if (results.length > 0) {
                        results.forEach((result) => {
                            documentItems.push({
                                Id: result.DocId,
                                FileName: result.Filename,
                                Url: result.LinkingUrl,
                                Title: result.Title,
                                PreviewImage: result.PictureThumbnailURL ? encodeURI(result.PictureThumbnailURL) : '',
                                Description: result.Description,
                                Path: encodeURI(result.Path),
                                ModifiedBY: result.ModifiedBy,
                                ModifiedDateTime: result.LastModifiedTime,
                                Type: result.FileExtension
                            });
                        });
                    }
                    return documentItems;
                });
            }
            else {
                return documentItems;
            }
        }
        catch (e) {
            return documentItems;
        }
    }

    public getDataFromQueryResults(data: any): any {
        if (data.PrimaryQueryResult.RelevantResults.Table.Rows) {
            var restResults = data.PrimaryQueryResult.RelevantResults.Table.Rows;
            var popularRestResults = new Array();
            if (restResults && restResults.results && restResults.results.length) {
                var results = restResults.results;
                for (var i = 0; i < results.length; i++) {
                    var propertyResults: any = {};
                    if (results[i].Cells && results[i].Cells.results && results[i].Cells.results.length) {
                        for (var j = 0; j < results[i].Cells.results.length; j++) {
                            propertyResults[results[i].Cells.results[j]["Key"]] = results[i].Cells.results[j]["Value"];
                        }
                    }

                    popularRestResults[i] = propertyResults;
                }
            }
            return popularRestResults;
        }
        else {
            return null;
        }
    }

    private getNestedChildren(arr, parent) {
        var out = [];
        for (var i in arr) {
            if (arr[i].parentId == parent.Id) {
                var Terms = this.getNestedChildren(arr, arr[i]);

                if (Terms.length) {
                    arr[i].Terms = Terms;
                }
                out.push(arr[i]);
            }
        }
        return out;
    }

    public async getNestsedChildTerms(terms) {
        var out = [];
        for (var i in terms) {
            if (terms[i].TermsCount > 0) {
                var childTerms = await terms[i].terms.get();
                var termsData = await this.getNestsedChildTerms(childTerms);
                out = out.concat(childTerms);
                if (termsData.length) {
                    out = out.concat(termsData);
                }
            }
        }
        return out;
    }

    public async GetTermstoreData(termstoreId: string, termGroupId: string, termsetId: string, termId: string): Promise<ITermstoreResult> {
        const taxonomy = new Session(this.ctx.pageContext.site.absoluteUrl);
        // var store = await taxonomy.termStores.getByName(PGPConstants.TaxonomyTermstoreName).get();
        // var set = store.getTermSetById(PGPConstants.TermsetGroupId);
        var store = await taxonomy.termStores.getById(termstoreId).get();
        var group, termset, data: any = [], childHierarchyValue, termsets, terms;
        childHierarchyValue = termId ? termId : (termsetId ? termsetId : (termGroupId ? termGroupId : ""));
        switch (childHierarchyValue) {
            case termId:
                var termDetails = await store.getTermById(termId).get();
                data = [termDetails];
                if (termDetails.TermsCount > 0) {
                    var childTerms = await termDetails.terms.get();
                    data = data.concat(childTerms);
                    var childterms = await this.getNestsedChildTerms(childTerms);
                    data = data.concat(childterms);
                }
                break;
            case termsetId:
                termset = store.getTermSetById(termsetId);
                data = await termset.terms.get();
                break;
            case termGroupId:
                group = await store.getTermGroupById(termGroupId).get();
                termsets = await group.termSets.get();
                var Promises = [];
                termsets.map((termsetItem) => {
                    Promises.push(termsetItem.terms.get());
                });
                await Promise.all(Promises).then((results) => {
                    results.map((result) => {
                        data = data.concat(result);
                    });
                });
                break;
        }
        var parentCategories = [];
        var taxonomyData = [];
        // data.map((term) => {
        //     var groupItem = term.PathOfTerm && term.PathOfTerm.split(";") ? term.PathOfTerm.split(";") : [];
        //     if (groupItem.length == 1 && categories.indexOf(groupItem[0]) == -1) {
        //         categories.push(term);
        //     }
        // });
        // categories.map((category) => {
        //     taxonomyData.push({
        //         Title: category.Name, TermDetails: category, Terms: data.filter((term) => {
        //             var groupItem = term.PathOfTerm.split(";");
        //             return groupItem.length > 1 && groupItem[0] == category.Name;
        //         })
        //     });
        // });

        data.map((term: any) => {
            if (term.Name == term.PathOfTerm) {
                term.isParent = true;
                parentCategories.push(term);
            }
            let parentLevel = (term.PathOfTerm.match(/;/g) || []).length;
            if (parentLevel > 0) {
                term.parentLevel = parentLevel;
                var parent = data.filter((v) => { return v.PathOfTerm == term.PathOfTerm.slice(0, term.PathOfTerm.lastIndexOf(";")); });
                var rootParent = data.filter((v) => { return v.PathOfTerm == term.PathOfTerm.slice(0, term.PathOfTerm.indexOf(";")); });
                term.parentId = parent.length > 0 ? parent[0].Id : "";
                term.rootParentId = rootParent.length > 0 ? rootParent[0].Id : "";
            }
        });

        parentCategories.map((parent) => {
            var thisCategoryInfo = data.filter((cat) => { return cat.rootParentId == parent.Id; }).sort((a: any, b: any) => {
                return (b.PathOfTerm.split(";").length) - (a.PathOfTerm.split(";").length);
            });
            thisCategoryInfo.sort((a: any, b: any) => {
                return (b.PathOfTerm.split(";").length) - (a.PathOfTerm.split(";").length);
            });
            taxonomyData.push({
                Title: parent.Name, TermDetails: parent, Terms: this.getNestedChildren(thisCategoryInfo, parent)
            });
        });
        return { GroupedData: taxonomyData, UnGroupedData: data };
    }

}