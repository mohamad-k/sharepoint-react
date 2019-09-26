import * as React from 'react';
import { sp } from '@pnp/sp';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';

import IDocumentType from '../models/IDocumentTypeModel';
import IStorageLocation, { StorageLocations } from '../models/IStorageLocationModel';
import IDocument, { CustomerContentType, DestructibleDocuments, RetrievedDocuments } from '../models/IDocumentModel';
import * as strings from 'ShareXarchivverwaltungWebPartStrings';
import styles from './ShareXarchivverwaltung.module.scss';

export interface IDocumentDialogProps {
  changeView: Function;
  showDialog: Function;
  documentTypes: IDocumentType[];
  documents: IDocument[];
  onFilter: Function;
  contentType: string;
}

export interface IDocumentDialogState {
  locationItems: IContextualMenuItem[];
  selectedCustomer: string;
  selectedContent: string;
  selectedTypes: string[];
  selectedLocation: IContextualMenuItem;
  selectedDate: Date | null;
}

export default class DocumentDialogS extends React.Component<IDocumentDialogProps, IDocumentDialogState> {
  public state: IDocumentDialogState = {
    locationItems: [],
    selectedCustomer: undefined,
    selectedContent: undefined,
    selectedLocation: undefined,
    selectedTypes: [],
    selectedDate: undefined
  };

  private _items: IContextualMenuItem[] = [
    {
      key: 'Content',
      name: 'Content',
      className: styles.commandBarItem,
      onRender: () => {
        return (
          <SearchBox
            placeholder={strings.content}
            onChange={(value: string) => {
              this.setState({ selectedContent: value }, () => {
                this._filter();
              });
            }}
          />
        );
      },
      onClick: (e: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
        e.preventDefault();
      }
    },
    {
      key: 'Location',
      name: 'Location',
      className: styles.commandBarItem,
      onRender: () => {
        return (
          <DefaultButton
            text={this.state.selectedLocation ? this.state.selectedLocation.name : strings.location}
            menuProps={{ items: this.state.locationItems }}
          />
        );
      },
      onClick: (e: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
        e.preventDefault();
      }
    },
    {
      key: 'Type',
      name: 'Type',
      className: styles.commandBarItem,
      onRender: () => {
        return (
          <Dropdown
            placeHolder={strings.documentType}
            options={this.props.documentTypes.map((type: IDocumentType) => ({ key: type.Id, text: type.Title }))}
            onChanged={this._onTypeSelected}
            multiSelect={true}
          />
        );
      },
      onClick: (e: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
        e.preventDefault();
      }
    },
    {
      key: 'Date',
      name: 'Date',
      className: styles.commandBarItem,
      onRender: () => {
        return (
          <DatePicker
            placeholder={strings.creationDate}
            allowTextInput={true}
            showWeekNumbers={true}
            firstWeekOfYear={1}
            firstDayOfWeek={DayOfWeek.Monday}
            showMonthPickerAsOverlay={true}
            formatDate={this._formatDate}
            onSelectDate={(date: Date | null | undefined) => {
              this.setState({ selectedDate: date }, () => {
                this._filter();
              });
            }}
            value={this.state.selectedDate}
          />
        );
      },
      onClick: (e: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
        e.preventDefault();
      }
    }
  ];

  constructor(props: IDocumentDialogProps) {
    super(props);

    this._onTypeSelected = this._onTypeSelected.bind(this);
  }

  public componentDidMount(): void {
    sp.web.lists
      .getByTitle('ArchivAblageorte')
      .items.select('*', 'StorageParent/Id', 'StorageParent/Title')
      .expand('StorageParent')
      .get()
      .then((locations: IStorageLocation[]) => {
        const items: IContextualMenuItem[] = this.mapItemsToContextualMenuItems(locations, null);
        this.setState({ locationItems: items });
      });
  }

  public render(): JSX.Element {
    return (
      <div className={styles.topbar}>
        <div className={styles.mtb5}>
          <DefaultButton
            onClick={() => {
              this.props.changeView('archivverwaltungview');
            }}
            iconProps={{ iconName: 'ChevronLeft' }}
            text={strings.archiveManagement}
          />
          {this._showAddBtn()}
        </div>
        <CommandBar className={`${styles.dFlex} ${styles.mtb5}`} isSearchBoxVisible={false} items={this._showCustomerField()} />
      </div>
    );
  }

  private _formatDate(date: Date): string {
    const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
    const deutsch = date.toLocaleDateString('de-DE', options);
    return deutsch;
  }

  private mapItemsToContextualMenuItems(locations: IStorageLocation[], parentId: number | null): IContextualMenuItem[] {
    const items: IContextualMenuItem[] = [];
    const filteredLocations = locations.filter((location: IStorageLocation) => location.StorageParentId === parentId);

    if (parentId === null) {
      const item: IContextualMenuItem = {
        key: 'clearLocationFilter',
        name: strings.deselect,
        iconProps: { iconName: 'ChromeClose' },
        onClick: () => {
          this.setState({ selectedLocation: undefined }, () => {
            this._filter();
          });
        }
      };

      items.push(item);
    }

    for (const location of filteredLocations) {
      const item: IContextualMenuItem = {
        key: location.Title,
        name: location.Title,
        data: location.StorageGroup,
        split: true,
        onClick: (e?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem) => {
          this.setState({ selectedLocation: item }, () => {
            this._filter();
          });
        }
      };

      const subItems = this.mapItemsToContextualMenuItems(locations, location.Id);

      if (subItems.length > 0) {
        item.subMenuProps = {
          items: subItems
        };
      }
      items.push(item);
    }
    return items;
  }

  private _onTypeSelected(option: IDropdownOption): void {
    if (option.selected) {
      this.setState({ selectedTypes: [...this.state.selectedTypes, option.text] }, () => {
        this._filter();
      });
    } else {
      const types = this.state.selectedTypes;
      types.splice(this.state.selectedTypes.indexOf(option.text), 1);
      this.setState({ selectedTypes: types }, () => {
        this._filter();
      });
    }
  }

  private _filter(): void {
    let filteredDocuments = undefined;

    if (this.state.selectedContent && this.state.selectedContent.length !== 0) {
      const docs = this.props.documents.filter((doc: IDocument) => {
        return doc.DocumentDescription.indexOf(this.state.selectedContent) !== -1;
      });
      filteredDocuments = docs;
    }

    if (this.state.selectedTypes && this.state.selectedTypes.length !== 0) {
      const documents = filteredDocuments ? filteredDocuments : this.props.documents;
      const docs = documents.filter((doc: IDocument) => {
        return this.state.selectedTypes.indexOf(doc.DocumentType.Title) !== -1;
      });
      filteredDocuments = docs;
    }

    if (this.state.selectedLocation) {
      const documents = filteredDocuments ? filteredDocuments : this.props.documents;
      const docs = documents.filter((doc: IDocument) => {
        switch (this.state.selectedLocation.data as string) {
          case StorageLocations.Building:
            return doc.StorageBuilding ? doc.StorageBuilding.Title === this.state.selectedLocation.name : null;
          case StorageLocations.Room:
            return doc.StorageRoom ? doc.StorageRoom.Title === this.state.selectedLocation.name : null;
          case StorageLocations.Rack:
            return doc.StorageRack ? doc.StorageRack.Title === this.state.selectedLocation.name : null;
          case StorageLocations.Shelf:
            return doc.StorageShelf ? doc.StorageShelf.Title === this.state.selectedLocation.name : null;
          case StorageLocations.Folder:
            return doc.StorageFolder ? doc.StorageFolder.Title === this.state.selectedLocation.name : null;
        }
      });
      filteredDocuments = docs;
    }
    if (this.state.selectedCustomer && this.state.selectedCustomer.length !== 0) {
      const documents = filteredDocuments ? filteredDocuments : this.props.documents;
      const docs = documents.filter((doc: IDocument) => {
        return doc.Customer.toUpperCase().indexOf(this.state.selectedCustomer.toUpperCase()) > -1;
      });
      filteredDocuments = docs;
    }
    if (this.state.selectedDate) {
      const documents = filteredDocuments ? filteredDocuments : this.props.documents;
      const docs = documents.filter((doc: IDocument) => {
        const date = this.props.contentType === DestructibleDocuments ? doc.DestructionDate : doc.EntryDate;
        return this.state.selectedDate.getTime() === new Date(date).getTime();
      });
      filteredDocuments = docs;
    }

    if (
      (!this.state.selectedContent || this.state.selectedContent.length === 0) &&
      (!this.state.selectedTypes || this.state.selectedTypes.length === 0) &&
      (!this.state.selectedLocation || !this.state.selectedLocation.hasOwnProperty('name')) &&
      (!this.state.selectedCustomer || this.state.selectedCustomer.length === 0) &&
      !this.state.selectedDate
    ) {
      filteredDocuments = this.props.documents;
    }
    this.props.onFilter(filteredDocuments);
  }

  private _showCustomerField(): IContextualMenuItem[] {
    const cItems = this._items.slice();

    if (this.props.contentType === CustomerContentType) {
      cItems.unshift({
        key: 'Search',
        name: 'Suche',
        className: styles.commandBarItem,
        onRender: () => {
          return (
            <SearchBox
              placeholder={strings.customerName}
              onChange={(value: string) => {
                this.setState({ selectedCustomer: value }, () => {
                  this._filter();
                });
              }}
            />
          );
        }
      });
    }
    return cItems;
  }

  private _showAddBtn(): JSX.Element {
    if (this.props.contentType === DestructibleDocuments || this.props.contentType === RetrievedDocuments) {
      return null;
    } else {
      return (
        <DefaultButton
          onClick={() => {
            this.props.showDialog();
          }}
          iconProps={{ iconName: 'Add' }}
          text={strings.archiveNewDocument}
        />
      );
    }
  }
}
