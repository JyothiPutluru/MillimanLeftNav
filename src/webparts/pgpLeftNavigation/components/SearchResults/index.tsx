import * as React from 'react';
import PGPConstants from '../pgp.constants';
import { IDocumentItem } from '../pgp.models';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from '../PgpLeftNavigation.module.scss';
export interface ISearchResultsProps {
    searchResults: IDocumentItem[];
}
export interface ISearchResultsState {
    activeItem: IDocumentItem;
    isItemExpanded: boolean;
    hoverItem: IDocumentItem;
}

export default class SearchResults extends React.Component<ISearchResultsProps, ISearchResultsState> {
    private categories: Array<any> = PGPConstants.Categories;

    public constructor(props: ISearchResultsProps) {
        super(props);
        this.state = {
            activeItem: null,
            isItemExpanded: false,
            hoverItem: null
        };
        this.expandSearchItem = this.expandSearchItem.bind(this);
        this.collapseSearchItem = this.collapseSearchItem.bind(this);
        this.showDownChevron = this.showDownChevron.bind(this);
    }

    private expandSearchItem(searchItem: IDocumentItem) {
        this.setState({ activeItem: searchItem, isItemExpanded: true });
    }

    private collapseSearchItem(searchItem: IDocumentItem) {
        this.setState({ activeItem: null, isItemExpanded: false, hoverItem: null });
    }

    private showDownChevron(searchItem: IDocumentItem) {
        if (!this.state.activeItem && (!this.state.hoverItem || (this.state.hoverItem && this.state.hoverItem.Id != searchItem.Id))) {
            this.setState({ hoverItem: searchItem, activeItem: null, isItemExpanded: false });
        }
    }

    public render(): React.ReactElement<ISearchResultsProps> {
        return (
            <div className={styles.pgp_search_results_pane}>
                {this.props.searchResults.map((searchDoc) => {
                    var breadcrumbFirstUrl = searchDoc.Path && searchDoc.Path.split('/').length > 0 ? searchDoc.Path.slice((searchDoc.Path.lastIndexOf('/')) + 1, searchDoc.Path.length - 1) : '';
                    var lastUrlPath = searchDoc.Path.slice(0, (searchDoc.Path.lastIndexOf('/')));
                    var breadcrumbSecondUrl = lastUrlPath && lastUrlPath.split('/').length > 0 ? lastUrlPath.slice((lastUrlPath.lastIndexOf('/')) + 1, lastUrlPath.length - 1) : '';
                    var iconClass = PGPConstants.FileIcons.filter((icon) => { return icon.format.toLowerCase() == searchDoc.Type.toLowerCase(); });
                    var iconClassName = iconClass.length > 0 ? iconClass[0].IconName : "";

                    return <div className={styles.doc_item_container} onMouseOver={this.showDownChevron.bind(this, searchDoc)}>
                        <div className={styles.doc_item_Icon}>
                            <Icon iconName={iconClassName} className="ms-IconExample" />
                        </div>
                        <div className={styles.doc_item_Content}>
                            <div className={styles.doc_item_title}><a href={searchDoc.Url}>{searchDoc.Title}</a></div>
                            <div className={styles.doc_item_sub_content}>
                                <div className={styles.doc_item_breadCrumb}><span>{decodeURI(breadcrumbSecondUrl)}</span> > <span>{decodeURI(breadcrumbFirstUrl)}</span></div>
                                <div className={styles.doc_item_sub_item}> Audience: <span> {searchDoc.ModifiedBY}</span></div>
                            </div>
                        </div>
                        {
                            <div className={styles.doc_item_hover_Content + (this.state.hoverItem && this.state.hoverItem.Id == searchDoc.Id ? ` ${styles.doc_item_hover_show_Content}` : ` ${styles.doc_item_hover_hide_Content}`)} >
                                {(!this.state.activeItem && !this.state.isItemExpanded && this.state.hoverItem && this.state.hoverItem.Id == searchDoc.Id) && <Icon iconName="DoubleChevronDown" className={styles.doc_item_Down_Chevron} onClick={this.expandSearchItem.bind(this, searchDoc)} />}
                                {this.state.activeItem && this.state.activeItem.Id == searchDoc.Id && this.state.isItemExpanded
                                    ? <div className={styles.doc_item_hover_expanded_Content}>
                                        <div className={styles.doc_item_hover_expanded_Content_pane}>
                                            <div className={styles.doc_item_hover_expanded_desc}>
                                                <span>{searchDoc.Description}</span>
                                            </div>
                                            <div className={styles.doc_item_hover_expanded_preview}>
                                                <img src={searchDoc.PreviewImage} alt="Preview" />
                                            </div>
                                        </div>
                                        <Icon iconName="DoubleChevronUp" className={styles.doc_item_Up_Chevron} onClick={this.collapseSearchItem.bind(this, searchDoc)} />
                                    </div>
                                    : <div />}
                            </div>
                        }
                    </div>;
                })}
            </div>
        );
    }
}
