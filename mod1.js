import IUser from './IUserModel';
import IDocumentType from './IDocumentTypeModel';
import IStorageLocation from './IStorageLocationModel';
import { taxonomy, ITermStore, ITermSet, ITerms, ITerm, ITermData, ITermStoreData, setItemMetaDataField } from '@pnp/sp-taxonomy';

export const CustomerContentType: string = '0x0100BBDF385CC6AB482E9B370B69B2755F86';
export const InstitutContentType: string = '0x01003FB3F85EB97D458BAB5B55CC86C752E0';
export const DestructibleDocuments: string = 'destructibleDocuments';
export const RetrievedDocuments: string = 'retrievedDocuments';

export default interface IDocument {
  Id: number;
  Title: string;
  Customer?: string;
  Docket: string;
  DocumentTypeId: number;
  DocumentType: IDocumentType;
  DocumentDescription: string;
  DataControllerId: number;
  DataController: IUser;
  Sensitivity: {id: number, TermGuid: string};
  EntryDate: string;
  DestructionDate: string;
  Indestructible: boolean;
  StorageBuildingId: number;
  StorageBuilding: IStorageLocation;
  StorageRoomId: number;
  StorageRoom: IStorageLocation;
  StorageRackId: number;
  StorageRack: IStorageLocation;
  StorageShelfId: number;
  StorageShelf: IStorageLocation;
  StorageFolderId: number;
  StorageFolder: IStorageLocation;
  HandedOut: boolean;
  HandedOutToId: number;
  HandedOutTo: IUser;
  HandedOutAt: string;
  HandedOutReason: string;
  Modified: string;
  Created: string;
  TaxCatchAll: {id: number, Term: string}[]
}


export const StorageLocations = {
    Building: 'Gebäude',
    Room: 'Raum',
    Rack: 'Regal',
    Shelf: 'Fach',
    Folder: 'Ordner'
  };
  
  export default interface IStorageLocation {
    Id?: number;
    Title: string;
    StorageParentId: number;
    StorageParent?: IStorageLocation;
    StorageGroup: string;
  }
  

  declare interface IShareXarchivverwaltungWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    DescriptionFieldLabel: string;
    months: string[];
    shortMonths: string[];
    days: string[];
    shortDays: string[];
    goToToday: string;
    editDocument: string;
    newDocument: string;
    docket: string;
    documentType: string;
    content: string;
    dataController: string;
    sensitivity: string;
    building: string;
    room: string;
    rack: string;
    shelf: string;
    folder: string;
    location: string;
    creationDate: string;
    destructionDate: string;
    archiveManagement: string;
    deselect: string;
    customerName: string;
    archiveNewDocument: string;
    inhouseDocuments: string;
    documentsToDestroy: string;
    checkedOutDocuments: string;
    customerDocuments: string;
    selectDate: string;
    clear: string;
    noDestruction: string;
    save: string;
    cancel: string;
    institutesDocuments: string;
    destructibleDocuments: string;
    retrievedDocuments: string;
    handOutDocument: string;
    handOutAt: string;
    handOutTo: string;
    handOutReason: string;
    handOut: string;
    documentOverview: string;
    checkInDocument: string;
    destroyDocument: string;
    warningMessages: {
      invalidDestructionDate: string;
      willBeDestroyed: string;
      emptyFields: string;
    };
    loading: string;
    Errors: {
      noCustomerDocuments: string;
      noDestructibleDocuments: string;
      noInstituteDocuments: string;
      noRetrievedDocuments: string;
    };
  }
  
  declare module 'ShareXarchivverwaltungWebPartStrings' {
    const strings: IShareXarchivverwaltungWebPartStrings;
    export = strings;
  }
  

  define([], function () {
    return {
      PropertyPaneDescription: 'Description',
      BasicGroupName: 'Group Name',
      DescriptionFieldLabel: 'Description Field',
      months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
      shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
      shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
      goToToday: 'Go to today',
      editDocument: 'Edit document',
      destroyDocument: 'Destroy document',
      newDocument: 'New document',
      docket: 'Docket',
      documentType: 'Document type',
      content: 'Content',
      dataController: 'Data Controller',
      sensitivity: 'Sensitivity',
      building: 'Building',
      room: 'Room',
      rack: 'Rack',
      shelf: 'Shelf',
      folder: 'Folder',
      location: 'Location',
      creationDate: 'Creation date',
      destructionDate: 'Destruction date',
      archiveManagement: 'Archive management',
      deselect: 'Deselect',
      customerName: 'Customer name',
      archiveNewDocument: 'Archive new document',
      inhouseDocuments: 'In-house documents',
      documentsToDestroy: 'Documents to destroy',
      checkedOutDocuments: 'Checked-out documents',
      customerDocuments: 'Customer documents',
      loading: 'loading...',
      selectDate: 'select a date...',
      clear: 'Clear',
      noDestruction: 'No Destruction',
      save: 'Save',
      cancel: 'Cancel',
      institutesDocuments: 'Institute\'s Documents',
      destructibleDocuments: 'Documents to be destroyed',
      customerDocuments: 'Customer Documents',
      retrievedDocuments: 'Retrieved Documents',
      handOutDocument: 'Handout Document',
      handOutAt: 'Handout Date',
      handOutTo: 'Handout to',
      handOutReason: 'Handout Reason',
      handOut: 'Handout',
      documentOverview: 'Document Overview',
      checkInDocument: 'Check in Document',
      warningMessages: {
        invalidDestructionDate: 'The Document is not destroyed because the destruction date has not yet been reached.',
        willBeDestroyed: 'Are you sure you want to destroy the Document?',
        emptyFields: 'Please fill in all Fields'
      },
      loading: 'Loading...',
      Errors: {
        noCustomerDocuments: 'There are no Customer Documents',
        noDestructibleDocuments: 'There are no Destructive Documents.',
        noInstituteDocuments:'There are no Documents owned by the Institute.',
        noRetrievedDocuments: ' There are no Handedout Documents'
      }
    };
  });
  