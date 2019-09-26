import * as React from 'react';
import { sp } from '@pnp/sp';
import { Dialog, DialogFooter, IDialogStyles } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DatePicker, DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { NormalPeoplePicker } from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { Label } from 'office-ui-fabric-react/lib/Label';

import * as strings from 'ShareXarchivverwaltungWebPartStrings';
import IDocumentModel, { CustomerContentType, DestructibleDocuments } from '../models/IDocumentModel';
import DocumentDialog from './DocumentDialog';
import styles from './ShareXarchivverwaltung.module.scss';
import IUser from '../models/IUserModel';
import IStorageLocation from '../models/IStorageLocationModel';

export interface IDocumentOverviewDialogProps {
  document: IDocumentModel;
  hidden: boolean;
  closeDocumentOverviewDialog: Function;
  reload: Function;
  contentType: string;
  customers: string[];
  locations: IStorageLocation[];
}

export interface IDocumentOverviewDialogState {
  document: IDocumentModel;
  isDialogHidden: boolean;
  isEditFormDialogHidden: boolean;
  isCheckoutDialogHidden: boolean;
  handedOutAt: Date;
  handedOutTo: number;
  handedOutReason: string;
  users: IUser[];
}

export default class DocumentOverviewDialog extends React.Component<IDocumentOverviewDialogProps, IDocumentOverviewDialogState> {
  public state: IDocumentOverviewDialogState = {
    document: undefined,
    isDialogHidden: this.props.hidden,
    isEditFormDialogHidden: true,
    isCheckoutDialogHidden: true,
    handedOutAt: new Date(),
    handedOutTo: undefined,
    handedOutReason: undefined,
    users: []
  };

  constructor(props: IDocumentOverviewDialogProps) {
    super(props);

    this._renderDialog = this._renderDialog.bind(this);
    this._renderCheckInOutButton = this._renderCheckInOutButton.bind(this);
    this._showEditDialog = this._showEditDialog.bind(this);
    this._hideEditDialog = this._hideEditDialog.bind(this);
    this._deleteDocument = this._deleteDocument.bind(this);
    this._checkDocumentOut = this._checkDocumentOut.bind(this);
    this._checkDocumentIn = this._checkDocumentIn.bind(this);
    this._onFilterChanged = this._onFilterChanged.bind(this);
    this._onUserSelected = this._onUserSelected.bind(this);
    this._renderOverview = this._renderOverview.bind(this);
  }

  public componentDidMount(): void {
    sp.web.siteUsers.get().then((users: IUser[]) => {
      this.setState({ users });
    });
  }

  public render(): JSX.Element {
    return (
      <div className={styles.documentOverviewDialog}>
        {this._renderDialog()}
        <DocumentDialog
          hidden={this.state.isEditFormDialogHidden}
          closeDialog={() => {
            this.setState({ isEditFormDialogHidden: true });
          }}
          contentType={this.props.contentType}
          document={this.props.document}
          customers={this.props.customers}
          reload={this.props.reload}
          closeDocumentOverview={this.props.closeDocumentOverviewDialog}
          locations={this.props.locations}
        />

        <Dialog
          getStyles={this.getStyles}
          hidden={this.state.isCheckoutDialogHidden}
          onDismiss={() => {
            this.setState({ isCheckoutDialogHidden: true });
          }}
          title={strings.handOutDocument}>
          <DatePicker
            label={strings.handOutAt}
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            firstWeekOfYear={1}
            showMonthPickerAsOverlay={true}
            placeholder={strings.selectDate}
            value={this.state.handedOutAt}
            onSelectDate={(date: Date) => {
              this.setState({ handedOutAt: date });
            }}
            formatDate={this._formatDate}
          />

          <Label>{strings.handOutTo}</Label>
          <NormalPeoplePicker
            onResolveSuggestions={this._onFilterChanged}
            key={'normal'}
            removeButtonAriaLabel={'Remove'}
            onItemSelected={this._onUserSelected}
          />

          <TextField
            label={strings.handOutReason}
            onChanged={(text: string) => {
              this.setState({ handedOutReason: text });
            }}
            multiline={true}
            rows={4}
            required={true}
          />

          <DialogFooter>
            <PrimaryButton onClick={this._checkDocumentOut} text={strings.handOut} />
            <DefaultButton
              onClick={() => {
                this.setState({ isCheckoutDialogHidden: true });
              }}
              text={strings.cancel}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _renderDialog(): JSX.Element {
    if (this.props.document) {
      return (
        <Dialog
          className={styles.documentOverviewDialog}
          getStyles={this.getStyles}
          hidden={this.props.hidden}
          onDismiss={() => {
            this.props.closeDocumentOverviewDialog();
          }}
          title={strings.documentOverview}>
          {this._renderOverview()}
          {this._isDestructible()}
        </Dialog>
      );
    }
  }

  private getStyles(): IDialogStyles {
    return {
      root: [],
      main: [
        {
          selectors: {
            ['@media (min-width: 400px)']: {
              maxWidth: '900px',
              minWidth: '800px'
            }
          }
        }
      ]
    };
  }

  private _renderCheckInOutButton(): JSX.Element {
    if (this.props.document && this.props.document.HandedOut) {
      return <PrimaryButton onClick={this._checkDocumentIn} text={strings.checkInDocument} />;
    }

    return (
      <PrimaryButton
        onClick={() => {
          this.setState({ isCheckoutDialogHidden: false });
        }}
        text={strings.handOutDocument}
      />
    );
  }

  private _renderOverview(): JSX.Element {
    return (
      <table className={styles.stripedTable}>
        <tr>
          <th>ID</th>
          <td>{this.props.document.Id}</td>
        </tr>
        {this.props.contentType === CustomerContentType && (
          <tr>
            <th>{strings.customerName}</th>
            <td>{this.props.document.Customer}</td>
          </tr>
        )}
        <tr>
          <th>{strings.content}</th>
          <td>{this.props.document.DocumentDescription}</td>
        </tr>
        <tr>
          <th>{strings.creationDate}</th>
          <td>{new Date(this.props.document.Created).toLocaleDateString('DE-de', { year: 'numeric', month: '2-digit', day: '2-digit' })}</td>
        </tr>
        <tr>
          <th>{strings.sensitivity}</th>
          <td>{this.props.document.TaxCatchAll[0].Term}</td>
        </tr>
        <tr>
          <th>{strings.noDestruction}</th>
          <td>{this.props.document.Indestructible ? 'Ja' : 'Nein'}</td>
        </tr>
        <tr>
          <th>{strings.documentType}</th>
          <td>{this.props.document.DocumentType.Title}</td>
        </tr>
      </table>
    );
  }

  private _checkDocumentOut(): void {
    sp.web.lists
      .getByTitle('ArchivDokumente')
      .items.getById(this.props.document.Id)
      .update({
        HandedOut: true,
        HandedOutAt: this.state.handedOutAt,
        HandedOutToId: this.state.handedOutTo,
        HandedOutReason: this.state.handedOutReason
      })
      .then(() => {
        this.setState({ isCheckoutDialogHidden: true }, () => {
          this.props.closeDocumentOverviewDialog();
        });
      });
  }

  private _checkDocumentIn(): void {
    sp.web.lists
      .getByTitle('ArchivDokumente')
      .items.getById(this.props.document.Id)
      .update({
        HandedOut: false,
        HandedOutAt: undefined,
        HandedOutToId: undefined,
        HandedOutReason: undefined
      })
      .then(() => {
        this.props.closeDocumentOverviewDialog();
        this.props.reload();
      });
  }

  private _showEditDialog(): void {
    this.setState({ isEditFormDialogHidden: false });
  }

  private _hideEditDialog(): void {
    this.setState({ isEditFormDialogHidden: true });
  }

  private _deleteDocument(): void {
    const date = new Date();
    date.setDate(date.getDate() - 1);
    if (this.props.document.DestructionDate === null) {
      alert('this Document is indestructible');
    } else if (new Date(this.props.document.DestructionDate).getTime() < date.getTime()) {
      const msg = confirm(strings.warningMessages.willBeDestroyed);
      if (msg) {
        sp.web.lists
          .getByTitle('ArchivDokumente')
          .items.getById(this.props.document.Id)
          .delete()
          .then(() => {
            this.props.closeDocumentOverviewDialog();
            this.props.reload();
          });
      }
    } else {
      alert(strings.warningMessages.invalidDestructionDate);
    }
  }

  private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[]): IPersonaProps[] {
    if (filterText) {
      let personas: IPersonaProps[] = this._filterPersonasByText(filterText);
      this._removeDuplicates(personas, currentPersonas);

      return personas;
    }

    return [];
  }

  private _filterPersonasByText(filterText: string): IPersonaProps[] {
    const personas: IPersonaProps[] = [];

    for (const user of this.state.users) {
      if (user.Title.toLowerCase().indexOf(filterText.toLowerCase()) !== -1) {
        personas.push({
          id: user.Id.toString(),
          text: user.Title
        });
      }
    }

    return personas;
  }

  private _onUserSelected(selectedUser: IPersonaProps): IPersonaProps {
    if (selectedUser) {
      this.setState({ handedOutTo: parseInt(selectedUser.id, 10) });
    }

    return selectedUser;
  }

  private _removeDuplicates(personas: IPersonaProps[], currentPersonas: IPersonaProps[]): void {
    for (const current of currentPersonas) {
      for (let i = 0; i < personas.length; i++) {
        if (personas[i].text === current.text) {
          personas.splice(i, 1);
        }
      }
    }
  }

  private _formatDate(date: Date): string {
    return date.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  private _isDestructible(): JSX.Element {
    if (this.props.contentType === DestructibleDocuments) {
      return (
        <DialogFooter>
          <PrimaryButton onClick={this._deleteDocument} text={strings.destroyDocument} />;
          <DefaultButton
            onClick={() => {
              this.props.closeDocumentOverviewDialog();
            }}
            text={strings.cancel}
          />
        </DialogFooter>
      );
    } else {
      return (
        <DialogFooter>
          <PrimaryButton onClick={this._showEditDialog} text={strings.editDocument} />
          {this._renderCheckInOutButton()}
          <PrimaryButton onClick={this._deleteDocument} text={strings.destroyDocument} />;
          <DefaultButton
            onClick={() => {
              this.props.closeDocumentOverviewDialog();
            }}
            text={strings.cancel}
          />
        </DialogFooter>
      );
    }
  }
}
