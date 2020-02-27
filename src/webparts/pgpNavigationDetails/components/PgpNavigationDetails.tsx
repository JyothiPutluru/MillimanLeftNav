import * as React from 'react';
import styles from './PgpNavigationDetails.module.scss';
import { IPgpNavigationDetailsProps } from './IPgpNavigationDetailsProps';
import { escape, isEqual } from '@microsoft/sp-lodash-subset';
import { mergeStyles, hiddenContentStyle, Dialog, Icon, DialogType, getId } from 'office-ui-fabric-react';

const screenReaderOnly = mergeStyles(hiddenContentStyle);

export interface IPgpNavigationDetailsState {
  term: any;
  hideDialog: boolean;
}

export default class PgpNavigationDetails extends React.Component<IPgpNavigationDetailsProps, IPgpNavigationDetailsState> {
  private _labelId: string = getId('dialogLabel');
  private _subTextId: string = getId('subTextLabel');
  public constructor(props: IPgpNavigationDetailsProps) {
    super(props);
    this.state = {
      term: null,
      hideDialog: true
    };
    this.getTermDetails = this.getTermDetails.bind(this);
    this._showDialog = this._showDialog.bind(this);
    this._closeDialog = this._closeDialog.bind(this);
  }
  
  public componentDidMount(){
    this.getTermDetails(this.props.termData);
  }

  public componentDidUpdate(prevprops:IPgpNavigationDetailsProps){
    if (!isEqual(this.props,prevprops)) {
      this.getTermDetails(this.props.termData);
    }
  }

  public getTermDetails(termData: any) {
    this.setState({ term: termData });
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  public render(): React.ReactElement<IPgpNavigationDetailsProps> {
    return (
      <div className={ styles.pgpNavigationDetails }>
        <div className={styles.container}>
          {this.state.term && this.state.term.Name
            ? <label id={this._labelId} className={screenReaderOnly}>
              {this.state.term.Name}
            </label>
            : null}

          { this.state.term && this.state.term.Description
            ? <label id={this._subTextId} className={screenReaderOnly}>
              {this.state.term.Description}
            </label>
            : null}

          {this.state.term && this.state.term.Name
            ?
            <div className={styles.termTitle} >{this.state.term.Name}
              {this.state.term.Description
                ? <span onClick={this._showDialog}><Icon className={styles.infoIcon} iconName="FullCircleMask"/><i className={styles.infoText}>i</i> </span>
                : null} </div>
            : <div />}

          {!this.state.hideDialog
            ? <Dialog
              hidden={this.state.hideDialog}
              onDismiss={this._closeDialog}
              dialogContentProps={{
                type: DialogType.normal,
                title: `${this.state.term ? this.state.term.Name:"-"}`,
                closeButtonAriaLabel: 'Close',
                subText: `${this.state.term ? this.state.term.Description:"-"}`
              }}
              modalProps={{
                titleAriaId: this._labelId,
                subtitleAriaId: this._subTextId,
                isBlocking: false
              }}
            >
            </Dialog>
            : null}
        </div>
      
        </div>
    );
  }
}
