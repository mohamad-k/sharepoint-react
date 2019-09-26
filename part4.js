import * as React from 'react';
import { sp } from '@pnp/sp';
import { DetailsList, DetailsListLayoutMode, IColumn, CheckboxVisibility } from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Link, LinkBase } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

import DocumentDialog from './DocumentDialog';
import IDocument, { CustomerContentType, InstitutContentType, DestructibleDocuments, RetrievedDocuments } from '../models/IDocumentModel';
import DocumentOverviewDialog from './DocumentOverviewDialog';
import IDocumentType from '../models/IDocumentTypeModel';
import Topbar from './Topbar';

import * as strings from 'ShareXarchivverwaltungWebPartStrings';
import styles from './ShareXarchivverwaltung.module.scss';
import IStorageLocation from '../models/IStorageLocationModel';

export interface IKundendokumenteViewProps {
  changeView: Function;
  contentType: string;
}

export interface IKundendokumenteViewState {
  filteredDocuments: IDocument[];
  columns: IColumn[];
  documents: IDocument[];
  documentTypes: IDocumentType[];
  documentOverWiewHidden: boolean;
  archiveDialogHidden: boolean;
  selectedDocument: IDocument;
  error: string;
  locations: IStorageLocation[];
}

export default class KundendokumenteView extends React.Component<IKundendokumenteViewProps, IKundendokumenteViewState> {
  public state: IKundendokumenteViewState = {
    filteredDocuments: [],
    columns: this._getColumns(),
    documents: [],
    documentTypes: [],
    documentOverWiewHidden: true,
    archiveDialogHidden: true,
    selectedDocument: undefined,
    error: undefined,
    locations: []
  };

  constructor(props: IKundendokumenteViewProps) {
    super(props);
    this._load = this._load.bind(this);
    this._onRenderItemColumn = this._onRenderItemColumn.bind(this);
    this._showArchiveDialog = this._showArchiveDialog.bind(this);
    this._onColumnClick = this._onColumnClick.bind(this);
    this._onItemInvoked = this._onItemInvoked.bind(this);
    this._onFilter = this._onFilter.bind(this);
  }

  public componentDidMount(): void {
    this._load();
  }

  public render(): JSX.Element {
    return (
      <div className={styles.documentView}>
        <p className="ms-fontSize-su">
          {this.props.contentType === InstitutContentType
            ? strings.inhouseDocuments
            : this.props.contentType === DestructibleDocuments
            ? strings.documentsToDestroy
            : this.props.contentType === RetrievedDocuments
            ? strings.checkedOutDocuments
            : strings.customerDocuments}
        </p>
        <div>
          <Topbar
            changeView={this.props.changeView}
            showDialog={this._showArchiveDialog}
            contentType={this.props.contentType}
            documentTypes={this.state.documentTypes}
            documents={this.state.documents}
            onFilter={this._onFilter}
          />
        </div>
        <DetailsList
          items={this.state.filteredDocuments}
          columns={this.state.columns}
          layoutMode={DetailsListLayoutMode.justified}
          onColumnHeaderClick={this._onColumnClick}
          onRenderItemColumn={this._onRenderItemColumn}
          checkboxVisibility={CheckboxVisibility.hidden}
        />
        {this._errorMessage()}
        {this._loadingSpinner()}
        <DocumentOverviewDialog
          document={this.state.selectedDocument}
          hidden={this.state.documentOverWiewHidden}
          contentType={this.props.contentType}
          closeDocumentOverviewDialog={() => {
            this.setState({ documentOverWiewHidden: true });
          }}
          customers={this.state.documents.map((doc: IDocument) => {
            return doc.Customer;
          })}
          reload={this._load}
          locations={this.state.locations}
        />
        <DocumentDialog
          hidden={this.state.archiveDialogHidden}
          closeDialog={() => {
            this.setState({ archiveDialogHidden: true });
          }}
          reload={this._load}
          contentType={this.props.contentType}
          customers={this.state.documents.map((doc: IDocument) => {
            return doc.Customer;
          })}
        />
      </div>
    );
  }

  private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    let { filteredDocuments, columns } = this.state;
    let isSortedDescending = column.isSortedDescending;
    // If we've sorted this column, flip it.
    if (column.isSorted) {
      isSortedDescending = !isSortedDescending;
    }

    // Sort the items.
    filteredDocuments = filteredDocuments!.concat([]).sort((a: IDocument, b: IDocument) => {
      let firstValue: string, secondValue: string;

      if (column.key === 'documentType') {
        firstValue = a.DocumentType.Title.toLowerCase();
        secondValue = b.DocumentType.Title.toLowerCase();
      }

      if (column.key === 'content') {
        firstValue = a.DocumentDescription.toLowerCase();
        secondValue = b.DocumentDescription.toLowerCase();
      }

      if (column.key === 'customer') {
        firstValue = a.Customer.toLowerCase();
        secondValue = b.Customer.toLowerCase();
      }

      if (column.key === 'building') {
        firstValue = a.StorageBuilding ? a.StorageBuilding.Title.toLowerCase() : '';
        secondValue = b.StorageBuilding ? b.StorageBuilding.Title.toLowerCase() : '';
      }

      if (column.key === 'room') {
        firstValue = a.StorageRoom ? a.StorageRoom.Title.toLowerCase() : '';
        secondValue = b.StorageRoom ? b.StorageRoom.Title.toLowerCase() : '';
      }

      if (column.key === 'rack') {
        firstValue = a.StorageRack ? a.StorageRack.Title.toLowerCase() : '';
        secondValue = b.StorageRack ? b.StorageRack.Title.toLowerCase() : '';
      }

      if (column.key === 'shelf') {
        firstValue = a.StorageShelf ? a.StorageShelf.Title.toLowerCase() : '';
        secondValue = b.StorageShelf ? b.StorageShelf.Title.toLowerCase() : '';
      }

      if (column.key === 'folder') {
        firstValue = a.StorageFolder ? a.StorageFolder.Title.toLowerCase() : '';
        secondValue = b.StorageFolder ? b.StorageFolder.Title.toLowerCase() : '';
      }

      if (column.key === 'date') {
        firstValue = this.props.contentType === DestructibleDocuments ? a.DestructionDate : a.EntryDate;
        secondValue = this.props.contentType === DestructibleDocuments ? b.DestructionDate : b.EntryDate;
      }

      if (isSortedDescending) {
        return firstValue > secondValue ? 1 : -1;
      } else {
        return firstValue > secondValue ? -1 : 1;
      }
    });
    // Reset the items and columns to match the state.
    this.setState({
      filteredDocuments: filteredDocuments,
      columns: columns!.map((col: IColumn) => {
        col.isSorted = col.key === column.key;

        if (col.isSorted) {
          col.isSortedDescending = isSortedDescending;
        }
        return col;
      })
    });
  };

  private _onRenderItemColumn(item: IDocument, index: number, column: IColumn): JSX.Element {
    const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
    switch (column.key) {
      case 'customer':
        return <p>{item.Customer}</p>;
      case 'documentType':
        return <p>{item.DocumentType.Title}</p>;
      case 'date':
        return (
          <p>
            {this.props.contentType === DestructibleDocuments
              ? new Date(item.DestructionDate).toLocaleDateString('de-DE', options)
              : new Date(item.EntryDate).toLocaleDateString('de-DE', options)}
          </p>
        );
      case 'building':
        return <p>{item.StorageBuilding.Title}</p>;
      case 'room':
        return <p>{item.StorageRoom.Title}</p>;
      case 'rack':
        return <p>{item.StorageRack.Title}</p>;
      case 'shelf':
        return <p>{item.StorageShelf.Title}</p>;
      case 'folder':
        return <p>{item.StorageFolder.Title}</p>;
      case 'content':
        return (
          <Link
            onClick={(e: React.MouseEvent<LinkBase>): void => {
              e.preventDefault();
              this._onItemInvoked(index);
            }}>
            {item.DocumentDescription}
          </Link>
        );
    }
  }

  private _onItemInvoked(index: number): void {
    this.setState({
      selectedDocument: this.state.filteredDocuments[index],
      documentOverWiewHidden: false
    });
  }

  private _showArchiveDialog(): void {
    this.setState({ archiveDialogHidden: false });
  }

  private _load(): void {
    const date = new Date();
    date.setDate(date.getDate() - 1);

    const select = sp.web.lists
      .getByTitle('ArchivDokumente')
      .items.select(
        '*',
        'DocumentType/Title',
        'DocumentType/StorageTime',
        'DataController/Id',
        'DataController/Title',
        'StorageBuilding/Id',
        'StorageBuilding/Title',
        'StorageFolder/Id',
        'StorageFolder/Title',
        'StorageRack/Id',
        'StorageRack/Title',
        'StorageRoom/Id',
        'StorageRoom/Title',
        'StorageShelf/Id',
        'StorageShelf/Title',
        'TaxCatchAll/Id',
        'TaxCatchAll/Term'
      )
      .expand('DocumentType', 'DataController', 'StorageBuilding', 'StorageFolder', 'StorageRack', 'StorageRoom', 'StorageShelf', 'TaxCatchAll');
    if (this.props.contentType === DestructibleDocuments) {
      select
        .filter(`DestructionDate lt datetime'${date.toISOString()}'`)
        .get()
        .then((vernichtendenDocumente: IDocument[]) => {
          if (vernichtendenDocumente.length === 0) {
            this.setState({ error: strings.Errors.noDestructibleDocuments, filteredDocuments: [] });
          } else {
            this.setState({ documents: vernichtendenDocumente, filteredDocuments: vernichtendenDocumente, error: undefined });
          }
        });
    }
    if (this.props.contentType === CustomerContentType || this.props.contentType === InstitutContentType) {
      select
        .filter(
          `startswith(ContentTypeId,'${this.props.contentType}') and (DestructionDate ge datetime'${date.toISOString()}' or Indestructible eq 1)`
        )
        .get()
        .then((documents: IDocument[]) => {
          if (documents.length === 0 && this.props.contentType === InstitutContentType) {
            this.setState({ error: strings.Errors.noInstituteDocuments });
          } else if (documents.length === 0) {
            this.setState({ error: strings.Errors.noCustomerDocuments, filteredDocuments: [] });
          } else {
            this.setState({ documents, filteredDocuments: documents, error: undefined });
          }
        });
    }
    if (this.props.contentType === RetrievedDocuments) {
      select
        .filter(`HandedOut eq 1`)
        .get()
        .then((documents: IDocument[]) => {
          if (documents.length === 0) {
            this.setState({ error: strings.Errors.noRetrievedDocuments, filteredDocuments: [] });
          } else {
            this.setState({ documents, filteredDocuments: documents, error: undefined });
          }
        });
    }

    sp.web.lists
      .getByTitle('ArchivDokumententypen')
      .items.getAll()
      .then((data: IDocumentType[]) => {
        this.setState({ documentTypes: data });
      });

    sp.web.lists
      .getByTitle('ArchivAblageorte')
      .items.getAll()
      .then((locations: IStorageLocation[]) => {
        this.setState({ locations });
      });
  }

  private _loadingSpinner(): JSX.Element {
    if (this.state.documents.length === 0 && !this.state.error) {
      return <Spinner label={strings.loading} />;
    }
  }

  private _errorMessage(): JSX.Element {
    if (this.state.error) {
      return (
        <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
          {this.state.error}
        </MessageBar>
      );
    }
  }

  private _onFilter(documents: IDocument[]): void {
    this.setState({ filteredDocuments: documents });
  }

  private _getColumns(): IColumn[] {
    const columns = [
      {
        key: 'content',
        name: strings.content,
        fieldName: 'DocumentDescription',
        minWidth: 300,
        isResizable: true
      },
      {
        key: 'documentType',
        name: strings.documentType,
        fieldName: 'DocumentTypeId',
        minWidth: 180,
        isResizable: true
      },
      {
        key: 'building',
        name: strings.building,
        fieldName: 'StorageBuildingId',
        minWidth: 100,
        isResizable: true
      },
      {
        key: 'room',
        name: strings.room,
        fieldName: 'StorageRoomId',
        minWidth: 100,
        isResizable: true
      },
      {
        key: 'rack',
        name: strings.rack,
        fieldName: 'StorageRackId',
        minWidth: 100,
        isResizable: true
      },
      {
        key: 'shelf',
        name: strings.shelf,
        fieldName: 'StorageShelfId',
        minWidth: 100,
        isResizable: true
      },
      {
        key: 'folder',
        name: strings.folder,
        fieldName: 'StorageFolderId',
        minWidth: 100,
        isResizable: true
      },
      {
        key: 'date',
        name: this.props.contentType === DestructibleDocuments ? strings.destructionDate : strings.creationDate,
        fieldName: this.props.contentType === DestructibleDocuments ? 'DestructionDate' : 'Created',
        minWidth: 150,
        isResizable: true
      }
    ];
    if (this.props.contentType === CustomerContentType) {
      columns.unshift({
        key: 'customer',
        name: strings.customerName,
        fieldName: 'Customer',
        minWidth: 300,
        isResizable: true
      });
    }
    return columns;
  }
}
