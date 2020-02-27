import * as React from 'react';
import PGPConstants from '../pgp.constants';
//import this.styles from '../Pgp.module.scss';
import * as Normalstyles from '../PgpLeftNavigation.module.scss';
import * as Treestyles from '../treeview.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IDocumentItem, ITermstoreData, ITerm } from '../pgp.models';

export interface ICategoriesState {
    showMoreCategories: string[];
    expandedCategories: string[];
}

export interface ICategoriesProps {
    termstoreCategories: ITermstoreData[];
    updateSelectedTermsets: (term: any) => any;
    selectedCategory: string;
    ungroupedCategoryData: ITerm[];
    showTitle: boolean;
    title: string;
    titleLinkUrl: string;
    isTreeView: boolean;
}

export default class Categories extends React.Component<ICategoriesProps, ICategoriesState> {
    private styles: any;
    public constructor(props: ICategoriesProps) {
        super(props);
        this.state = {
            showMoreCategories: [],
            expandedCategories: []
        };
        this.ToggleShowMoreCategories = this.ToggleShowMoreCategories.bind(this);
        this.ToggleExpandedCategories = this.ToggleExpandedCategories.bind(this);
        this.getNestedChildrenUI = this.getNestedChildrenUI.bind(this);
        this.styles = this.props.isTreeView ? Treestyles.default : Normalstyles.default;
    }

    public componentWillReceiveProps(nxtProps: ICategoriesProps) {
        var stateObj: ICategoriesState = this.state;
        if (JSON.stringify(this.props.termstoreCategories) != JSON.stringify(nxtProps.termstoreCategories) && nxtProps.termstoreCategories) {
            stateObj.expandedCategories = nxtProps.termstoreCategories.map((t) => { return t.TermDetails.Id; });
            // this.setState({ expandedCategories: nxtProps.termstoreCategories.map((term) => { return term.TermDetails.Id; }) });
        }
        if (this.props.selectedCategory != nxtProps.selectedCategory || this.props.ungroupedCategoryData !=nxtProps.ungroupedCategoryData) {
            var term: any = nxtProps.ungroupedCategoryData.filter((t) => { return t.Id == nxtProps.selectedCategory; });
            if (term.length > 0) {
                stateObj.expandedCategories.push(term[0].Id);
                while (term[0].parentId) {
                    term = nxtProps.ungroupedCategoryData.filter((t) => { return t.Id == term[0].parentId; });
                    if (term.length > 0) {
                        stateObj.expandedCategories.push(term[0].Id);
                    }
                }
                stateObj.expandedCategories.push(term[0].Id);
            }
        }

        if (JSON.stringify(this.state) != JSON.stringify(stateObj)) {
            this.setState(stateObj);
        }
    }

    private ToggleShowMoreCategories(termId: string) {
        var showMoreCategories = this.state.showMoreCategories;
        if (this.state.showMoreCategories.indexOf(termId) == -1) {
            showMoreCategories.push(termId);
        } else {
            showMoreCategories.splice(this.state.showMoreCategories.indexOf(termId), 1);
        }
        this.setState({ showMoreCategories: showMoreCategories });
    }

    private ToggleExpandedCategories(termId: string) {
        var expandedCategories = this.state.expandedCategories;
        if (this.state.expandedCategories.indexOf(termId) == -1) {
            expandedCategories.push(termId);
        } else {
            expandedCategories.splice(this.state.expandedCategories.indexOf(termId), 1);
        }
        this.setState({ expandedCategories: expandedCategories });
    }

    private getNestedChildrenUI(category) {
        var out;
        if (category.Terms && category.Terms.length > 0) {
            out = (<div className={this.styles.pgp_side_category_body}>
                {category.Terms.map((subCategory) => {
                    return <div>
                        <div className={((subCategory.Terms && subCategory.Terms.length > 0 && this.props.isTreeView) ? this.styles.pgp_side_category_header : this.styles.pgp_side_category_content) + ` ${this.props.selectedCategory.indexOf(subCategory.Id) != -1 ? this.styles.pgp_side_category_selected : ""}`}>
                            {subCategory.Terms && subCategory.Terms.length > 0
                                ? <Icon iconName={(this.state.expandedCategories.indexOf(subCategory.Id) != -1) ? 'ChevronDownSmall' : 'ChevronRightSmall'}
                                    onClick={this.ToggleExpandedCategories.bind(this, subCategory.Id)}
                                />
                                : null}
                            <span onClick={this.props.updateSelectedTermsets.bind(this, subCategory)}>
                                {subCategory.Name}
                            </span>
                        </div>
                        {(this.state.expandedCategories.indexOf((subCategory.Id)) >= 0)
                            ? this.getNestedChildrenUI(subCategory)
                            : null}
                    </div>;
                })
                }
            </div >);
        }
        return out;
    }

    public render(): React.ReactElement<ICategoriesProps> {
        var context = this;
        if (this.props.isTreeView) {
            return (
                <div className={this.styles.pgp_category_pane}>
                    {this.props.showTitle&&this.props.ungroupedCategoryData.length
                        ?
                        <div className={`${this.styles.pgp_side_category_overview} ${!this.props.selectedCategory?this.styles.pgp_side_category_selected:""} `}><a href={this.props.titleLinkUrl ? this.props.titleLinkUrl : location.href.split("#")[0]}>{this.props.title}</a></div>
                        : null
                    }
                    {this.props.termstoreCategories.map((category) => {
                        var categories: any = JSON.stringify(category.Terms);
                        categories = JSON.parse(categories);
                        return <div className={``}>
                            <div className={this.styles.pgp_side_category_header + ` ${this.props.selectedCategory.indexOf(category.TermDetails.Id) != -1 ? this.styles.pgp_side_category_selected : ""}`} >
                                {category.Terms.length > 0 ? <Icon iconName={this.state.expandedCategories.indexOf(category.TermDetails.Id) != -1 ? 'ChevronDownSmall' : 'ChevronRightSmall'}
                                    onClick={this.ToggleExpandedCategories.bind(this, category.TermDetails.Id)}
                                /> : null}
                                <span onClick={this.props.updateSelectedTermsets.bind(this, category.TermDetails)}> {category.Title}</span>
                            </div>
                            {this.state.expandedCategories.indexOf(category.TermDetails.Id) >= 0 ?
                                this.getNestedChildrenUI(category)
                                : null}
                        </div>;
                    })}

                </div>
            );
        } else {
            return (
                <div className={this.styles.pgp_category_pane}>
                    {this.props.showTitle
                        ?
                        <div className="">
                            <div className={this.styles.pgp_side_category_overview} ><a href={this.props.titleLinkUrl ? this.props.titleLinkUrl : "#"}>{this.props.title}</a></div>
                        </div>
                        : null
                    }
                    {this.props.termstoreCategories.map((category) => {
                        var categories: any = JSON.stringify(category.Terms);
                        categories = JSON.parse(categories);
                        var splicedCategories = categories.splice(0, PGPConstants.noOfTermsetsToShow);
                        return <div className={``}>
                            <div className={this.styles.pgp_side_category_header + ` ${this.props.selectedCategory.indexOf(category.TermDetails.Id) != -1 ? this.styles.pgp_side_category_selected : ""}`} onClick={this.props.updateSelectedTermsets.bind(this, category.TermDetails)}>
                                {category.Title}
                            </div>
                            <div>
                                {category.Terms.length > PGPConstants.noOfTermsetsToShow && context.state.showMoreCategories.indexOf(category.TermDetails.Id) == -1
                                    ?
                                    splicedCategories.map((subCategory) => {
                                        return <div className={this.styles.pgp_side_category_content + ` ${this.props.selectedCategory.indexOf(subCategory.Id) != -1 ? this.styles.pgp_side_category_selected : ""}`} onClick={this.props.updateSelectedTermsets.bind(this, subCategory)}> {subCategory.Name}</div>;
                                    })
                                    : category.Terms.map((subCategory) => {
                                        return <div className={this.styles.pgp_side_category_content + ` ${this.props.selectedCategory.indexOf(subCategory.Id) != -1 ? this.styles.pgp_side_category_selected : ""}`} onClick={this.props.updateSelectedTermsets.bind(this, subCategory)}> {subCategory.Name}</div>;
                                    })
                                }

                                <div className={`${category.Terms.length > PGPConstants.noOfTermsetsToShow && context.state.showMoreCategories.indexOf(category.TermDetails.Id) == -1 ? this.styles.pgp_side_category_showMore : this.styles.pgp_side_category_hide}`} onClick={this.ToggleShowMoreCategories.bind(this, category.TermDetails.Id)}>
                                    Show More
                                </div>

                                <div className={`${category.Terms.length > PGPConstants.noOfTermsetsToShow && context.state.showMoreCategories.indexOf(category.TermDetails.Id) != -1 ? this.styles.pgp_side_category_showMore : this.styles.pgp_side_category_hide}`} onClick={this.ToggleShowMoreCategories.bind(this, category.TermDetails.Id)}>
                                    Show Less
                                </div>

                            </div>
                        </div>;
                    })}

                </div>
            );
        }
    }
}
