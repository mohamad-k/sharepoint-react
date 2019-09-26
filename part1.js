import * as React from 'react';
import { sp } from '@pnp/sp';
import { DocumentCard, DocumentCardPreview, DocumentCardTitle, IDocumentCardPreviewProps } from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';

import IDocument, { InstitutContentType, CustomerContentType, DestructibleDocuments, RetrievedDocuments } from '../models/IDocumentModel';
import styles from './ShareXarchivverwaltung.module.scss';
import * as strings from 'ShareXarchivverwaltungWebPartStrings';

export interface IArchiveDashboardViewState {
  destructibleDocuments: IDocument[];
  retrievedDocuments: IDocument[];
}

export interface IArchiveDashboardViewProps {
  changeView: Function;
}

export default class ArchiveDashboardView extends React.Component<IArchiveDashboardViewProps, IArchiveDashboardViewState> {
  public state: IArchiveDashboardViewState = {
    destructibleDocuments: [],
    retrievedDocuments: []
  };

  private previewProps: IDocumentCardPreviewProps = {
    previewImages: [
      {
        name: 'Revenue stream proposal fiscal year 2016 version02.pptx',
        previewImageSrc: 'https://via.placeholder.com/200',
        imageFit: ImageFit.cover,
        height: 150
      }
    ]
  };

  public componentDidMount(): void {
    const date = new Date();
    date.setDate(date.getDate() - 1);
    sp.web.lists
      .getByTitle('ArchivDokumente')
      .items.filter(`DestructionDate lt datetime'${date.toISOString()}'`)
      .select('ID')
      .get()
      .then((destructibleDocuments: IDocument[]) => {
        this.setState({ destructibleDocuments });
      });
    sp.web.lists
      .getByTitle('ArchivDokumente')
      .items.filter(`HandedOut eq 1`)
      .select('ID')
      .get()
      .then((retrievedDocuments: IDocument[]) => {
        this.setState({ retrievedDocuments });
      });
  }

  public render(): JSX.Element {
    return (
      <div className={`${styles.dFlex} ${styles.jcCenter}`}>
        <DocumentCard
          className={`${styles.ml5} ${styles.mr5} ms-textAlignCenter`}
          onClick={() => {
            this.props.changeView('dokumenteview', InstitutContentType);
          }}>
          <DocumentCardPreview {...this.previewProps} />
          <DocumentCardTitle showAsSecondaryTitle={true} title={strings.institutesDocuments} />
        </DocumentCard>

        <DocumentCard
          className={`${styles.ml5} ${styles.mr5} ms-textAlignCenter`}
          onClick={() => {
            this.props.changeView('dokumenteview', CustomerContentType);
          }}>
          <DocumentCardPreview {...this.previewProps} />
          <DocumentCardTitle showAsSecondaryTitle={true} title={strings.customerDocuments} />
        </DocumentCard>

        <DocumentCard
          className={`${styles.ml5} ${styles.mr5} ms-textAlignCenter`}
          onClick={() => {
            this.props.changeView('dokumenteview', RetrievedDocuments);
          }}>
          <DocumentCardPreview {...this.previewProps} />
          <div className={styles.badge}>{this.state.retrievedDocuments.length}</div>
          <DocumentCardTitle showAsSecondaryTitle={true} title={strings.retrievedDocuments} />
        </DocumentCard>

        <DocumentCard
          className={`${styles.ml5} ${styles.mr5} ms-textAlignCenter`}
          onClick={() => {
            this.props.changeView('dokumenteview', DestructibleDocuments);
          }}>
          <DocumentCardPreview {...this.previewProps} />
          <div className={styles.badge}>{this.state.destructibleDocuments.length}</div>
          <DocumentCardTitle showAsSecondaryTitle={true} title={strings.destructibleDocuments} />
        </DocumentCard>
      </div>
    );
  }
}
