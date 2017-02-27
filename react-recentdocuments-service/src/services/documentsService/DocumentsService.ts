import { IDocument } from './IDocument';
import { IDocumentsService } from './IDocumentsService';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { Promise } from 'es6-promise';

export class DocumentsService implements IDocumentsService {
    public static readonly serviceKey: ServiceKey<IDocumentsService> = ServiceKey.create<IDocumentsService>('contoso:DocumentsService', DocumentsService);    
    private static documents: IDocument[] = [
        {
            title: 'Proposal for Jacksonville Expansion Ad Campaign',
            url: 'https://mod156629-my.sharepoint.com/personal/miriamg_mod156629_onmicrosoft_com/_layouts/15/WopiFrame.aspx?sourcedoc=%7BCBF65183-0378-485B-AB67-791E0FC81D72%7D&file=Jacksonville%20Ad%20Campaign%20(draft).docx&action=view&DefaultItemOpen=1',
            imageUrl: 'https://nam.delve.office.com/mt/v3/documents/preview?siteid=923a6ce1-7b67-4bd0-a59f-89d37f233804&webid=c12486eb-661c-46c7-baba-073a8a45b610&uniqueid=d50f5428-979f-428c-bddd-244e1342e954&docid=266210838&token=&path=&spWebUrl=&clienttype=DelveWeb',
            iconUrl: '',
            activity: {
                title: 'Modified, January 25 2017',
                actorName: 'Miriam Graham',
                actorImageUrl: 'https://mod156629-my.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=miriamg@mod156629.onmicrosoft.com&size=L'
            }
        },
        {
            title: 'Customer Feedback for ZT1000',
            url: 'https://mod156629-my.sharepoint.com/personal/miriamg_mod156629_onmicrosoft_com/_layouts/15/WopiFrame.aspx?sourcedoc=%7B5449CE24-BFB7-442E-843D-E0C86CEB71CC%7D&file=Customer%20Feedback%20for%20ZT1000.pptx&action=view&DefaultItemOpen=1',
            imageUrl: 'https://nam.delve.office.com/mt/v3/documents/preview?siteid=923a6ce1-7b67-4bd0-a59f-89d37f233804&webid=c12486eb-661c-46c7-baba-073a8a45b610&uniqueid=638c29da-dfbb-43c8-aad7-69277a93e409&docid=266238609&token=&path=&spWebUrl=&clienttype=DelveWeb',
            iconUrl: '',
            activity: {
                title: 'Modified, January 23 2017',
                actorName: 'Miriam Graham',
                actorImageUrl: 'https://mod156629-my.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=miriamg@mod156629.onmicrosoft.com&size=L'
            }
        },
        {
            title: 'Asia Q3 Marketing Overview',
            url: 'https://mod156629-my.sharepoint.com/personal/alexw_mod156629_onmicrosoft_com/_layouts/15/WopiFrame.aspx?sourcedoc=%7BFD077A94-AB7D-45F9-A810-36229E518A94%7D&file=Asia%20Q3%20Marketing%20Overview%20Beta.pptx&action=view&DefaultItemOpen=1',
            imageUrl: 'https://nam.delve.office.com/mt/v3/documents/preview?siteid=923a6ce1-7b67-4bd0-a59f-89d37f233804&webid=c12486eb-661c-46c7-baba-073a8a45b610&uniqueid=1f22c3a7-be84-4873-85e0-aafd06a082b6&docid=266227443&token=&path=&spWebUrl=&clienttype=DelveWeb',
            iconUrl: '',
            activity: {
                title: 'Modified, January 23 2017',
                actorName: 'Alex Wilber',
                actorImageUrl: 'https://mod156629-my.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=alexw@mod156629.onmicrosoft.com&size=L'
            }
        },
        {
            title: 'Trey Research Business Development Plan',
            url: 'https://mod156629.sharepoint.com/sites/contoso/Resources/Document%20Center/_layouts/15/WopiFrame.aspx?sourcedoc=%7B743A6C44-D3F8-4ECC-A1B7-EA9844911314%7D&file=Trey%20Research%20Business%20Development%20Plan.pptx&action=view&DefaultItemOpen=1',
            imageUrl: 'https://nam.delve.office.com/mt/v3/documents/preview?siteid=923a6ce1-7b67-4bd0-a59f-89d37f233804&webid=c12486eb-661c-46c7-baba-073a8a45b610&uniqueid=8b3d16c2-4dc3-4a13-bd69-e4b503f1a2a2&docid=265998810&token=&path=&spWebUrl=&clienttype=DelveWeb',
            iconUrl: '',
            activity: {
                title: 'Modified, January 15 2017',
                actorName: 'Alex Wilber',
                actorImageUrl: 'https://mod156629-my.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=alexw@mod156629.onmicrosoft.com&size=L'
            }
        },
        {
            title: 'XT1000 Marketing Analysis',
            url: 'https://mod156629-my.sharepoint.com/personal/henriettam_mod156629_onmicrosoft_com/_layouts/15/WopiFrame.aspx?sourcedoc=%7BA8B9F935-E5A1-47AD-B052-D5ED30E682AB%7D&file=XT1000%20Marketing%20Analysis.pptx&action=view&DefaultItemOpen=1',
            imageUrl: 'https://nam.delve.office.com/mt/v3/documents/preview?siteid=b1d447c2-592a-4b2f-95dc-33c1d3dbabc1&webid=237a3f3f-59a4-46e8-b0a8-6effd78bd327&uniqueid=fb23ec89-bb37-44f9-8813-d2dfd447340d&docid=17592963535263&token=&path=&spWebUrl=&clienttype=DelveWeb',
            iconUrl: '',
            activity: {
                title: 'Modified, December 15 2016',
                actorName: 'Henrietta Mueller',
                actorImageUrl: 'https://mod156629-my.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=henriettam@mod156629.onmicrosoft.com&size=L'
            }
        }
    ];
    private loadedData: IDocument[];
    private loadingData: boolean;

    constructor(serviceScope: ServiceScope) {
        console.warn('Creating DocumentsService');
    }

    public getRecentDocument(): Promise<IDocument> {
        return new Promise<IDocument>((resolve: (document: IDocument) => void, reject: (error: any) => void): void => {
            this.ensureRecentDocuments()
                .then((recentDocuments: IDocument[]): void => {
                    resolve(recentDocuments[0]);
                });
        });
    }

    public getRecentDocuments(startFrom: number = 0): Promise<IDocument[]> {
        return new Promise<IDocument[]>((resolve: (documents: IDocument[]) => void, reject: (error: any) => void): void => {
            this.ensureRecentDocuments()
                .then((recentDocuments: IDocument[]): void => {
                    resolve(recentDocuments.slice(startFrom, startFrom + 3));
                });
        });
    }

    private ensureRecentDocuments(): Promise<IDocument[]> {
        return new Promise<IDocument[]>((resolve: (recentDocuments: IDocument[]) => void, reject: (error: any) => void): void => {
            if (this.loadedData) {
                console.info('Data already loaded');
                resolve(this.loadedData);
                return;
            }

            if (this.loadingData) {
                console.info('Data already being loaded. Waiting to complete...');
                window.setTimeout((): void => {
                    this.ensureRecentDocuments()
                        .then((recentDocuments: IDocument[]): void => {
                            resolve(recentDocuments);
                        });
                }, 100);
            }
            else {
                this.loadingData = true;
                console.warn('Retrieving recent documents...');
                window.setTimeout((): void => {
                    this.loadedData = DocumentsService.documents;
                    this.loadingData = false;
                    resolve(this.loadedData);
                }, DocumentsService.randomDelay());
            }
        });
    }

    private static randomDelay(): number {
        const min: number = 200;
        const max: number = 1000;
        return Math.random() * (max - min) + min;
    }
}