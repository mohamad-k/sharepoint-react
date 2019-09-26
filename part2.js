import * as React from 'react';
import { sp } from '@pnp/sp';
import { Dialog, DialogFooter, IDialogStyles } from 'office-ui-fabric-react/lib/Dialog';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { NormalPeoplePicker, ValidationState } from 'office-ui-fabric-react/lib/Pickers';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { taxonomy, ITermData, setItemMetaDataField } from '@pnp/sp-taxonomy';

import * as strings from 'ShareXarchivverwaltungWebPartStrings';
import styles from './ShareXarchivverwaltung.module.scss';
import IDocument, { CustomerContentType } from '../models/IDocumentModel';
import IDocumentType from '../models/IDocumentTypeModel';
import IUser from '../models/IUserModel';
import IStorageLocation, { StorageLocations } from '../models/IStorageLocationModel';

const DayPickerStrings: IDatePickerStrings = {
  months: strings.months,
  shortMonths: strings.shortMonths,
  days: strings.days,
  shortDays: strings.shortDays,
  goToToday: strings.goToToday
};

export interface IDocumentDialogProps {
  hidden: boolean;
  closeDialog: Function;
  reload?: Function;
  document?: IDocument;
  contentType: string;
  customers?: string[];
  closeDocumentOverview?: Function;
  locations?: IStorageLocation[];
}

export interface IDocumentDialogState {
  users: IPersonaProps[];
  fileNumber: string;
  content: string;
  sensitivityClass: string;
  dateOfEntry: null | Date;
  noDestruction: boolean;
  documentTypes: IDocumentType[];
  sensitivityClasses: ITermData[];
  customer?: IPersonaProps;
  dataController: IPersonaProps;
  documentTypeId: number;
  destructionDate: Date | null;
  storageTime: number;
  buildings: IStorageLocation[];
  selectedBuilding: IDropdownOption;
  rooms: IStorageLocation[];
  selectedRoom: IDropdownOption;
  racks: IStorageLocation[];
  selectedRack: IDropdownOption;
  shelfs: IStorageLocation[];
  selectedShelf: IDropdownOption;
  folders: IStorageLocation[];
  selectedFolder: IDropdownOption;
  customersNames: IPersonaProps[];
}

export default class DocumentDialog extends React.Component<IDocumentDialogProps, IDocumentDialogState> {
  public state: IDocumentDialogState = {
    users: [],
    fileNumber: undefined,
    content: undefined,
    sensitivityClass: undefined,
    dateOfEntry: this._getDate(),
    noDestruction: false,
    documentTypes: [],
    sensitivityClasses: [],
    customer: undefined,
    dataController: undefined,
    documentTypeId: undefined,
    destructionDate: undefined,
    storageTime: undefined,
    buildings: [],
    selectedBuilding: undefined,
    rooms: [],
    selectedRoom: undefined,
    racks: [],
    selectedRack: undefined,
    shelfs: [],
    selectedShelf: undefined,
    folders: [],
    selectedFolder: undefined,
    customersNames: []
  };

  constructor(props: IDocumentDialogProps) {
    super(props);
    this._onEntryDate = this._onEntryDate.bind(this);
    this._onSliderChange = this._onSliderChange.bind(this);
    this._closeDialog = this._closeDialog.bind(this);
    this._saveDocument = this._saveDocument.bind(this);
    this._onCustomerSelected = this._onCustomerSelected.bind(this);
    this._onControllerChange = this._onControllerChange.bind(this);
    this._setDestructionDate = this._setDestructionDate.bind(this);
    this._onStorageLocationChange = this._onStorageLocationChange.bind(this);
    this._onCustomerInputChanged = this._onCustomerInputChanged.bind(this);
  }

  public componentDidMount(): void {
    sp.web.siteUsers.get().then((users: IUser[]): void => {
      let usersNames = [];
      users.map((user: IUser) => {
        usersNames.push({ id: user.Id, text: user.Title });
      });
      this.setState({ users: usersNames });
    });
    sp.web.lists
      .getByTitle('ArchivDokumententypen')
      .items.getAll()
      .then((documenttypes: IDocumentType[]) => {
        this.setState({ documentTypes: documenttypes });
      });
    taxonomy.termStores
      .getByName('Managed Metadata Service')
      .groups.getByName('Site Collection - sharexpert.de-sites-Mohammed')
      .termSets.getByName('Sensitivität')
      .terms.get()
      .then((sensitivityClasses: ITermData[]) => {
        this.setState({ sensitivityClasses });
      });
    sp.web.lists
      .getByTitle('ArchivAblageorte')
      .items.filter(`StorageGroup eq '${encodeURI(StorageLocations.Building)}'`)
      .get()
      .then((buildings: IStorageLocation[]) => {
        this.setState({
          buildings: buildings
        });
      });

    taxonomy.termStores
      .getByName('Managed Metadata Service').groups
      .get()
      .then(res => console.log(res));
  }

  public componentWillReceiveProps(nextprops: IDocumentDialogProps): void {
    if (nextprops.document) {
      this.setState({
        fileNumber: nextprops.document.Docket,
        content: nextprops.document.DocumentDescription,
        sensitivityClass: `/Guid(${nextprops.document.Sensitivity.TermGuid})/`,
        noDestruction: nextprops.document.Indestructible,
        documentTypeId: nextprops.document.DocumentTypeId,
        dataController: { id: nextprops.document.DataController.Id.toString(), text: nextprops.document.DataController.Title },
        customer: { text: nextprops.document.Customer },
        dateOfEntry: new Date(nextprops.document.EntryDate),
        destructionDate: nextprops.document.Indestructible ? undefined : new Date(nextprops.document.DestructionDate),
        storageTime: nextprops.document.DocumentType.StorageTime,
        selectedBuilding: nextprops.document.StorageBuilding
          ? {
              key: nextprops.document.StorageBuilding.Id,
              text: nextprops.document.StorageBuilding.Title
            }
          : null,
        selectedRoom: nextprops.document.StorageRoom
          ? {
              key: nextprops.document.StorageRoom.Id,
              text: nextprops.document.StorageRoom.Title
            }
          : null,
        selectedRack: nextprops.document.StorageRack
          ? {
              key: nextprops.document.StorageRack.Id,
              text: nextprops.document.StorageRack.Title
            }
          : null,
        selectedShelf: nextprops.document.StorageShelf
          ? {
              key: nextprops.document.StorageShelf.Id,
              text: nextprops.document.StorageShelf.Title
            }
          : null,
        selectedFolder: nextprops.document.StorageFolder
          ? {
              key: nextprops.document.StorageFolder.Id,
              text: nextprops.document.StorageFolder.Title
            }
          : null,
        buildings: nextprops.locations.filter((location: IStorageLocation) => {
          return location.StorageGroup === StorageLocations.Building;
        }),
        rooms: nextprops.locations.filter((location: IStorageLocation) => {
          if (this.state.selectedBuilding) return this.state.selectedBuilding.key === location.StorageParentId;
        }),
        racks: nextprops.locations.filter((location: IStorageLocation) => {
          if (this.state.selectedRoom) return this.state.selectedRoom.key === location.StorageParentId;
        }),
        shelfs: nextprops.locations.filter((location: IStorageLocation) => {
          if (this.state.selectedRack) return this.state.selectedRack.key === location.StorageParentId;
        }),
        folders: nextprops.locations.filter((location: IStorageLocation) => {
          if (this.state.selectedShelf) return this.state.selectedShelf.key === location.StorageParentId;
        })
      });
    }
  }

  public render(): JSX.Element {
    return (
      <Dialog
        className={styles.documentDialog}
        getStyles={this._getStyles}
        hidden={this.props.hidden}
        onDismiss={this._closeDialog}
        title={this.props.document ? strings.editDocument : strings.newDocument}>
        {this._showCustomerName()}
        <div className={styles.dFlex}>
          <TextField
            className={styles.flexGrow1}
            label={strings.docket}
            required={true}
            onChanged={(val: string) => {
              this.setState({ fileNumber: val });
            }}
            value={this.state.fileNumber}
          />
          <Dropdown
            label={strings.documentType}
            selectedKey={this.state.documentTypeId}
            onChanged={(val: { key: number; text: string; time: number }) => {
              this.setState({ documentTypeId: val.key as number, storageTime: val.time }, this._setDestructionDate);
            }}
            options={this.state.documentTypes.map((type: IDocumentType) => ({ key: type.Id, text: type.Title, time: type.StorageTime }))}
            required={true}
          />
        </div>
        <TextField
          className={styles.mlr5}
          label={strings.content}
          multiline={true}
          rows={4}
          required={true}
          onChanged={(val: string) => {
            this.setState({ content: val });
          }}
          value={this.state.content}
        />
        <div className={styles.dFlex}>
          <div className={styles.flexGrow1}>
            <Label required={true}>{strings.dataController}:</Label>
            <NormalPeoplePicker
              onResolveSuggestions={this._onFilterController}
              selectedItems={this.state.dataController ? [this.state.dataController] : []}
              onChange={this._onControllerChange}
              itemLimit={1}
            />
          </div>
          <Dropdown
            label={strings.sensitivity}
            selectedKey={this.state.sensitivityClass}
            onChanged={(val: { key: string; text: string }) => {
              this.setState({ sensitivityClass: val.key });
            }}
            options={this.state.sensitivityClasses.map((klasse: ITermData) => ({ key: klasse.Id, text: klasse.Name }))}
            required={true}
          />
        </div>
        <div className={styles.dFlex}>
          <Dropdown
            label={strings.building}
            onChanged={this._onStorageLocationChange}
            selectedKey={this.state.selectedBuilding ? this.state.selectedBuilding.key : undefined}
            options={this.state.buildings.map((building: IStorageLocation) => ({
              key: building.Id,
              text: building.Title,
              data: building.StorageGroup
            }))}
          />
          <Dropdown
            label={strings.room}
            onChanged={this._onStorageLocationChange}
            selectedKey={this.state.selectedRoom ? this.state.selectedRoom.key : undefined}
            options={this.state.rooms.map((room: IStorageLocation) => ({ key: room.Id, text: room.Title, data: room.StorageGroup }))}
            isDisabled={this.state.rooms.length === 0}
          />
          <Dropdown
            label={strings.rack}
            onChanged={this._onStorageLocationChange}
            selectedKey={this.state.selectedRack ? this.state.selectedRack.key : undefined}
            options={this.state.racks.map((rack: IStorageLocation) => ({ key: rack.Id, text: rack.Title, data: rack.StorageGroup }))}
            isDisabled={this.state.racks.length === 0}
          />
          <Dropdown
            label={strings.shelf}
            onChanged={this._onStorageLocationChange}
            selectedKey={this.state.selectedShelf ? this.state.selectedShelf.key : undefined}
            options={this.state.shelfs.map((shelf: IStorageLocation) => ({ key: shelf.Id, text: shelf.Title, data: shelf.StorageGroup }))}
            isDisabled={this.state.shelfs.length === 0}
          />
          <Dropdown
            label={strings.folder}
            onChanged={this._onStorageLocationChange}
            selectedKey={this.state.selectedFolder ? this.state.selectedFolder.key : undefined}
            options={this.state.folders.map((folder: IStorageLocation) => ({ key: folder.Id, text: folder.Title, data: folder.StorageGroup }))}
            isDisabled={this.state.folders.length === 0}
          />
        </div>
        <div className={styles.dFlex}>
          <DatePicker
            className={styles.flexGrow1}
            onSelectDate={this._onEntryDate}
            firstDayOfWeek={DayOfWeek.Monday}
            allowTextInput={true}
            strings={DayPickerStrings}
            showWeekNumbers={true}
            firstWeekOfYear={1}
            showMonthPickerAsOverlay={true}
            placeholder={strings.selectDate}
            value={this.state.dateOfEntry}
            isRequired={true}
            formatDate={this._formatDate}
            label={strings.creationDate}
          />
          <DefaultButton
            className={`${styles.alignSelfEnd}`}
            onClick={() => {
              this.setState({ dateOfEntry: undefined });
            }}
            text={strings.clear}
          />
        </div>
        <div className={styles.dFlex}>
          <DatePicker
            className={styles.flexGrow1}
            label={strings.destructionDate}
            value={this.state.destructionDate}
            formatDate={this._formatDate}
            disabled={true}
          />
          <Toggle
            className={`${styles.dFlex} ${styles.flexColumn}`}
            onChanged={this._onSliderChange}
            checked={this.state.noDestruction}
            label={strings.noDestruction}
          />
        </div>

        <DialogFooter>
          <PrimaryButton onClick={this._saveDocument} text={strings.save} />
          <DefaultButton onClick={this._closeDialog} text={strings.cancel} />
        </DialogFooter>
      </Dialog>
    );
  }

  private _onEntryDate(date: Date): void {
    this.setState({ dateOfEntry: date }, this._setDestructionDate);
  }

  private _onSliderChange(isChecked: boolean): void {
    if (isChecked) {
      this.setState({ noDestruction: true, destructionDate: undefined });
    } else {
      this.setState({ noDestruction: false }, this._setDestructionDate);
    }
  }

  private _getStyles(): IDialogStyles {
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

  private _onCustomerSelected(customers: IPersonaProps[]): void {
    this.setState({ customer: customers[0] });
  }

  private _onCustomerInputChanged(input: string): ValidationState {
    return ValidationState.valid;
  }

  private _onControllerChange(users: IPersonaProps[]): void {
    this.setState({ dataController: users[0] });
  }

  private _onFilterController = (
    filterText: string,
    currentPersonas: IPersonaProps[],
    limitResults?: number
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = this._filterControllerByText(filterText);
      return (filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas);
    } else {
      return [];
    }
  };

  private _onFilterCustomer = (
    filterText: string,
    currentPersonas: IPersonaProps[],
    limitResults?: number
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = this._filetCustomersByText(filterText);
      return (filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas);
    } else {
      return [];
    }
  };

  private _filetCustomersByText(filterText: string): IPersonaProps[] {
    let uniq = {};
    let customers: IPersonaProps[] = [];
    this.props.customers.map((customer: string) => customers.push({ primaryText: customer }));
    let removeDuplicatedCustomers = this.props.customers
      ? customers.filter((obj: IPersonaProps) => !uniq[obj.primaryText] && (uniq[obj.primaryText] = true))
      : [];
    return removeDuplicatedCustomers.filter((user: IPersonaProps) => this._doesTextStartWith(user.primaryText, filterText));
  }

  private _filterControllerByText(filterText: string): IPersonaProps[] {
    return this.state.users.filter((user: IPersonaProps) => this._doesTextStartWith(user.text, filterText));
  }

  private _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) > -1;
  }

  private _saveDocument(): void {
    let {
      customer,
      fileNumber,
      documentTypeId,
      content,
      dataController,
      sensitivityClass,
      noDestruction,
      dateOfEntry,
      destructionDate,
      selectedBuilding,
      selectedRoom,
      selectedRack,
      selectedShelf,
      selectedFolder
    } = this.state;

    let condition =
      this.props.contentType === CustomerContentType
        ? fileNumber && DocumentType && content && sensitivityClass && dateOfEntry && dataController && customer
        : fileNumber && DocumentType && content && sensitivityClass && dateOfEntry && dataController;

    if (condition) {
      const setItemMetaData = async document => {
        setItemMetaDataField(
          document.item,
          'Sensitivity',
          await taxonomy
            .getDefaultSiteCollectionTermStore()
            .getTermById(sensitivityClass)
            .get()
        );
      };
      if (this.props.document) {
        sp.web.lists
          .getByTitle('ArchivDokumente')
          .items.getById(this.props.document.Id)
          .update({
            Customer: customer.primaryText,
            Docket: fileNumber,
            DocumentTypeId: documentTypeId,
            DocumentDescription: content,
            DataControllerId: dataController.id,
            Modified: new Date(),
            EntryDate: dateOfEntry,
            Indestructible: noDestruction,
            DestructionDate: noDestruction ? null : destructionDate,
            StorageBuildingId: selectedBuilding ? selectedBuilding.key : null,
            StorageFolderId: selectedFolder ? selectedFolder.key : null,
            StorageRackId: selectedRack ? selectedRack.key : null,
            StorageRoomId: selectedRoom ? selectedRoom.key : null,
            StorageShelfId: selectedShelf ? selectedShelf.key : null
          })
          .then(setItemMetaData)
          .then(() => {
            this.props.closeDocumentOverview();
          })
          .then(() => {
            this._closeDialog();
            this.props.reload();
          });
      } else {
        sp.web.lists
          .getByTitle('ArchivDokumente')
          .items.add({
            Customer: this.props.contentType === CustomerContentType ? customer.primaryText : null,
            Docket: fileNumber,
            DocumentTypeId: documentTypeId,
            DocumentDescription: content,
            DataControllerId: parseInt(dataController.id, 10),
            EntryDate: dateOfEntry,
            DestructionDate: noDestruction ? undefined : destructionDate,
            Indestructible: noDestruction,
            ContentTypeId: this.props.contentType,
            StorageBuildingId: selectedBuilding ? selectedBuilding.key : null,
            StorageFolderId: selectedFolder ? selectedFolder.key : null,
            StorageRackId: selectedRack ? selectedRack.key : null,
            StorageRoomId: selectedRoom ? selectedRoom.key : null,
            StorageShelfId: selectedShelf ? selectedShelf.key : null
          })
          .then(setItemMetaData)
          .then(() => {
            this._closeDialog();
            this.props.reload();
          });
      }
    } else {
      alert(strings.warningMessages.emptyFields);
    }
  }

  private _closeDialog(): void {
    this.setState(
      {
        customer: undefined,
        fileNumber: undefined,
        content: undefined,
        sensitivityClass: undefined,
        noDestruction: false,
        dataController: undefined,
        destructionDate: undefined,
        selectedBuilding: undefined,
        selectedRoom: undefined,
        rooms: [],
        selectedRack: undefined,
        racks: [],
        selectedShelf: undefined,
        shelfs: [],
        selectedFolder: undefined,
        folders: [],
        documentTypeId: undefined,
        dateOfEntry: this._getDate()
      },
      () => this.props.closeDialog()
    );
  }

  private _formatDate(date: Date): string {
    const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
    const deutsch = date.toLocaleDateString('de-DE', options);
    return deutsch;
  }

  private _setDestructionDate(): void {
    if (this.state.dateOfEntry && this.state.documentTypeId && !this.state.noDestruction) {
      this.setState({ destructionDate: new Date(`December 31, ${this.state.dateOfEntry.getFullYear() + this.state.storageTime}`) });
    }
  }

  private _showCustomerName(): JSX.Element {
    if (this.props.contentType === CustomerContentType) {
      return (
        <div className={styles.mlr5}>
          <Label required={true}>{strings.customerName}</Label>
          <NormalPeoplePicker
            onResolveSuggestions={this._onFilterCustomer}
            selectedItems={this.state.customer ? [this.state.customer] : []}
            onChange={this._onCustomerSelected}
            onValidateInput={this._onCustomerInputChanged}
            itemLimit={1}
          />
        </div>
      );
    }
  }

  private _onStorageLocationChange(item: IDropdownOption): void {
    switch (item.data) {
      case StorageLocations.Building:
        this.setState({ selectedBuilding: item, rooms: [], racks: [], shelfs: [], folders: [] });
        break;
      case StorageLocations.Room:
        this.setState({ selectedRoom: item, racks: [], shelfs: [], folders: [] });
        break;
      case StorageLocations.Rack:
        this.setState({ selectedRack: item, shelfs: [], folders: [] });
        break;
      case StorageLocations.Shelf:
        this.setState({ selectedShelf: item, folders: [] });
        break;
      case StorageLocations.Folder:
        this.setState({ selectedFolder: item });
        break;
    }

    sp.web.lists
      .getByTitle('ArchivAblageorte')
      .items.filter(`StorageParentId eq ${item.key}`)
      .get()
      .then((locations: IStorageLocation[]) => {
        if (locations.length > 0) {
          switch (locations[0].StorageGroup) {
            case StorageLocations.Room:
              this.setState({ rooms: locations });
              break;
            case StorageLocations.Rack:
              this.setState({ racks: locations });
              break;
            case StorageLocations.Shelf:
              this.setState({ shelfs: locations });
              break;
            case StorageLocations.Folder:
              this.setState({ folders: locations });
              break;
          }
        }
      });
  }

  private _getDate(): Date {
    const date: Date = new Date();
    date.setHours(0, 0, 0, 0);
    return date;
  }
}
