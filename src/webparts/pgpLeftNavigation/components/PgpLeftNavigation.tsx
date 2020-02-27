import * as React from 'react';
import styles from './PgpLeftNavigation.module.scss';
import { IPgpLeftNavigationProps } from './IPgpLeftNavigationProps';
import { IRefinementFilter } from '../../../models/ISearchResult';
import { IFilterData, ISearchRecievedProperties, IDocumentItem, ITermstoreData, ITerm } from './pgp.models';
import PGPSerive, { IPGPSerive, ITermstoreResult } from './pgp.svc';
import Categories from '../components/Categories/index';

export interface IPgpLeftNavigationState {
	searchResults: IDocumentItem[];
	termstoreCategories: ITermstoreData[];
	unGroupedData: ITerm[];
	selectedCategory: string;
	previousSelectedCategory: string;
}

export interface IHashObject {
	[key: string]: string;
	Refiners: string;
	SearchText: string;
	RequiredProperty: string;
}

export default class PgpLeftNavigation extends React.Component<IPgpLeftNavigationProps, IPgpLeftNavigationState> {

	private PGPSerive: IPGPSerive;
	public constructor(props: IPgpLeftNavigationProps) {
		super(props);
		this.searchResults = this.searchResults.bind(this);
		this.updateSelectedTermsets = this.updateSelectedTermsets.bind(this);
		this.updatedSelectedItem = this.updatedSelectedItem.bind(this);
		this.PGPSerive = new PGPSerive(this.props.context);
		this.state = {
			termstoreCategories: [],
			unGroupedData: [],
			selectedCategory: "",
			searchResults: [],
			previousSelectedCategory: "",
		};
	}

	public componentDidMount() {
		let reactHandler = this;
		window.addEventListener('hashchange', () => {
			reactHandler.updatedSelectedItem();
		}, false);
		
		if (this.props.taxonomyTermstoreId && (this.props.taxonomyTermGroupId || this.props.taxonomyTermsetId || this.props.taxonomyTermId)) {
			this.PGPSerive.GetTermstoreData(this.props.taxonomyTermstoreId, this.props.taxonomyTermGroupId, this.props.taxonomyTermsetId, this.props.taxonomyTermId).then((data: ITermstoreResult) => {
				this.setState({ termstoreCategories: data.GroupedData, unGroupedData: data.UnGroupedData });
				this.updatedSelectedItem();
				sessionStorage.setItem("unGroupedTermsData", JSON.stringify(this.state.unGroupedData));
				sessionStorage.setItem("requiredProperty", this.props.mappedField);
			});
		}
	}

	public componentDidUpdate(prevprops: IPgpLeftNavigationProps) {
		if (this.state.unGroupedData.length) {
			let refiners: IRefinementFilter[] = [];
			let selectedTermId = this.state.selectedCategory ? this.state.selectedCategory : this.getHashKeyValue(this.props.mappedField) ? `/Guid(${this.getHashKeyValue(this.props.mappedField)})/` : "";
			refiners = this.getHashKeyValue("Refiners") ? JSON.parse(this.getHashKeyValue("Refiners")) : [];
			if (JSON.stringify(refiners) == JSON.stringify(prevprops.refinerSourceData.selectedFilters)) {
				refiners = this.props.refinerSourceData.selectedFilters;
			}
			let changedProperties: ISearchRecievedProperties = {
				Refiners: refiners,
				SearchBoxWebPart: (this.props.searchQuery && selectedTermId === this.state.previousSelectedCategory) || (this.props.searchQuery != prevprops.searchQuery) ? this.props.searchQuery : this.getHashKeyValue("SearchText"),
				QueryKeywords: this.props.searchQuery ? this.props.searchQuery + " " + this.formatTextFilter() : this.formatTextFilter(),
				TermData: this.getSetectedTerm(),
				RefinerSourceData: {
					refinerConfiguration: this.props.refinerSourceData.refinerConfiguration,
					selectedFilters: refiners
				}
			};

			if (this.props.refinerSourceData) {
				this.updateHashOnBulkModification({
					SearchText: changedProperties.SearchBoxWebPart,
					Refiners: JSON.stringify(refiners),
					[this.props.mappedField]: changedProperties.TermData ? `${changedProperties.TermData.Id.slice(changedProperties.TermData.Id.indexOf("(") + 1, changedProperties.TermData.Id.indexOf(")"))}` : "",
					RequiredProperty: this.props.mappedField,
				});
			}
		}
	}

	public render(): React.ReactElement<IPgpLeftNavigationProps> {
		return (
			<div className={styles.pgpLeftNavigation}>
				<div className={styles.container}>
					{(this.props.taxonomyTermstoreId && (this.props.taxonomyTermGroupId || this.props.taxonomyTermsetId || this.props.taxonomyTermId))
						?
						<Categories termstoreCategories={this.state.termstoreCategories}
							updateSelectedTermsets={this.updateSelectedTermsets}
							selectedCategory={this.state.selectedCategory}
							showTitle={this.props.showTitle}
							title={this.props.title}
							ungroupedCategoryData={this.state.unGroupedData}
							titleLinkUrl={this.validateSearchUrl(this.props.titleLinkUrl) ? this.props.titleLinkUrl : location.href.split("#")[0]}
							isTreeView={this.props.isTreeView}
						/>
						: <div style={{ padding: '30px' }}>Please configure the webpart</div>
					}
					<style>{`
				           div[class^=markdown]{
				            font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
				            font-size: 14px;
				            line-height: 1.4285;
				            color: #333;
				           }
				          `}
					</style>
				</div>
			</div>
		);
	}

	private formatTextFilter() {
		let queryJoining: string = "";
		if (this.getHashKeyValue(this.props.mappedField) && this.getHashKeyValue("SearchText")) {
			queryJoining = " AND ";
		}
		return this.getHashKeyValue(this.props.mappedField) ? `${queryJoining}${this.props.mappedField}:${this.getHashKeyValue(this.props.mappedField)}` : "";
	}

	private getDataFromURLHash(): ISearchRecievedProperties {
		return {
			Refiners: this.getHashKeyValue("Refiners"),
			SearchBoxWebPart: this.getHashKeyValue("SearchText"),
			QueryKeywords: this.getHashKeyValue("SearchText") + this.formatTextFilter(),
			TermData: this.getSetectedTerm(),
			RefinerSourceData: {
				refinerConfiguration: this.props.refinerSourceData.refinerConfiguration,// : this.defaultRefinerConfig,
				selectedFilters: this.getHashKeyValue("Refiners") ? JSON.parse(this.getHashKeyValue("Refiners")) : []
			}
		};
	}

	protected validateSearchUrl(value: string): boolean {
		var regex = new RegExp(/^(http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/)?[a-z0-9]+([\-\.]{1}[a-z0-9]+)*\.[a-z]{2,5}(:[0-9]{1,5})?(\/.*)?$/);
		value = encodeURI(value);
		var isValid = regex.test(value);
		return isValid;
	}

    /*
    * Function called when URL hash changes.
    */
	private getSearchData(changedProperties: ISearchRecievedProperties) {
		if (changedProperties.TermData) {
			changedProperties.QueryKeywords = this.getHashKeyValue("SearchText") + this.formatTextFilter();
		}
		else if (this.props.showTitle) {
			changedProperties.SearchBoxWebPart = "";
			changedProperties.TermData = {
				parentId: {},
				Id: this.props.title,
				Name: this.props.title,
				Owner: this.props.title,
				PathOfTerm: this.props.title,
				Description: this.props.title
			};
			if (!this.validateSearchUrl(this.props.titleLinkUrl)) {
				changedProperties.QueryKeywords = this.props.titleLinkUrl;
			}
		}
		if (!this.state.selectedCategory) {
			changedProperties.SearchBoxWebPart = "";
		}
		this.props.transmitSearchProperties(changedProperties);
	}

	private getHashKeyValue(hashKey: string): any {
		let hashKeyValue;
		let hashStr = decodeURIComponent(window.location.hash);
		if (hashStr) {
			var decodedHashStr = decodeURIComponent(atob(hashStr.substring(1)));
			let hashStrObj = JSON.parse(decodedHashStr.replace(/&quot;/g, '"'));
			hashKeyValue = hashStrObj[hashKey] ? hashStrObj[hashKey] : "";
		}
		return hashKeyValue;
	}

	private updateHashOnBulkModification(hashObject: IHashObject) {
		const pageUrl = decodeURIComponent(window.location.href);
		let hashStr = decodeURIComponent(window.location.hash);
		let newHashStr = JSON.stringify(hashObject);
		let encodedHashStr = btoa(encodeURIComponent(newHashStr));
		let updatedUrl: string = hashStr ? pageUrl.replace(`${hashStr}`, `#${encodedHashStr}`) : pageUrl.concat(`#${encodedHashStr}`);
		if (decodeURIComponent(window.location.href) != updatedUrl)
			window.location.href = updatedUrl;
	}

	public searchResults(searchText: string) {
		this.PGPSerive.GetSearchResults(searchText).then((results) => {
			this.setState({ searchResults: results });
		});
	}

	public updateSelectedTermsets(term) {
		let _previousSelectedCategory: string = this.state.selectedCategory != term.Id ? this.state.selectedCategory : "";
		this.setState({ selectedCategory: term.Id, previousSelectedCategory: _previousSelectedCategory });
		var selectedTerm = `${term.Id.slice(term.Id.indexOf("(") + 1, term.Id.indexOf(")"))}`;
		this.updateHashOnBulkModification({
			SearchText: "",
			Refiners: "[]",
			[this.props.mappedField]: selectedTerm,
			RequiredProperty: this.props.mappedField,
		});
	}

	private getSetectedTerm() {
		let termId: string = "";
		let hashStr = window.location.hash;
		try {
			if (hashStr) {
				var decodedHashStr = atob(hashStr.substring(1));
				let hashStrObj = JSON.parse(decodeURIComponent(decodedHashStr).replace(/&quot;/g, '"'));
				if (hashStrObj[this.props.mappedField]) {
					termId = `/Guid(${hashStrObj[this.props.mappedField]})/`;
					var term = this.state.unGroupedData.filter((dataItem) => { return dataItem.Id == termId; });
					if (term.length > 0) {
						return term[0];
					}
				}
			}
		}
		catch (error) {
			return null;
		}
	}

	private updatedSelectedItem() {
		let selectedTermData: ITerm = this.getSetectedTerm();
		if (selectedTermData) {
			let _previousSelectedCategory: string = this.state.selectedCategory != selectedTermData.Id ? this.state.selectedCategory : "";
			this.setState({
				previousSelectedCategory: _previousSelectedCategory,
				selectedCategory: selectedTermData.Id,
			});
		}
		else if (!(JSON.parse(decodeURIComponent(atob(location.hash.substring(1))))[JSON.parse(decodeURIComponent(atob(location.hash.substring(1)))).RequiredProperty] || this.props.showTitle)) {
			let selectedTermGuid = this.state.termstoreCategories[2].TermDetails.Id;
			this.setState({
				selectedCategory: selectedTermGuid
			});
			var selectedTerm = `${selectedTermGuid.slice(selectedTermGuid.indexOf("(") + 1, selectedTermGuid.indexOf(")"))}`;
			this.updateHashOnBulkModification({
				SearchText: "",
				Refiners: "[]",
				[this.props.mappedField]: selectedTerm,
				RequiredProperty: this.props.mappedField,
			});
		}
		else {
			this.setState({
				selectedCategory: "",
			});
		}
		this.getSearchData(this.getDataFromURLHash());
	}
}
