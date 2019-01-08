import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as pnp from 'sp-pnp-js';
import { ITag } from 'office-ui-fabric-react/lib/Pickers';

export interface ISitePagesService {
    _pageContext: PageContext;
    getSitePages: (rowLimit: number, tags: ITag[]) => Promise<any[]>;
    getRoles: () => Promise<any[]>;
    getDepartments: () => Promise<any[]>;
}

export class SitePagesService implements ISitePagesService {
    public static readonly serviceKey: ServiceKey<ISitePagesService> = ServiceKey.create<ISitePagesService>('ImageGallery:SitePagesService', SitePagesService);
    public _pageContext: PageContext;    

    constructor(serviceScope: ServiceScope) {
        

        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
        });
    }

    public getSitePages(rowLimit: number, tags: ITag[]): Promise<any[]> {
        const listName = 'Site Pages';
        const xml = `<View>
                        <ViewFields>
                            <FieldRef Name='ID' />
                            <FieldRef Name='Title' />
                            <FieldRef Name='ImageLink' />
                            <FieldRef Name='NavigationURL' />
                        </ViewFields>
                        <Query>
                            <OrderBy>
                                <FieldRef Name='SortOrder' />
                            </OrderBy>
                        </Query>
                        <RowLimit>` + rowLimit + `</RowLimit>
                    </View>`;

        const q: any = {
            ViewXml: xml,
        };

        //return this._ensureList(listName).then((list) => {
        //    if (list) {
                //,

                let filterText = '';

                tags.forEach(tag => {
                    if (tag.key.indexOf('Dept-',0) >= 0) {
                        const tagId = tag.key.replace('Dept-','');
                        filterText += "Department/Id eq " + tagId;
                    }

                    if (tag.key.indexOf('Role-',0) >= 0) {
                        const tagId = tag.key.replace('Role-','');
                        filterText += "Relevant_x0020_Roles/Id eq " + tagId;
                    }
                });

                if (tags.length > 0 && filterText === '') {
                    // Hide all items
                    filterText = 'Id eq 0';
                }

                return pnp.sp.web.lists
                    .getByTitle(listName)
                    .items
                    .select('FileLeafRef','Title','UniqueId','Modified','Relevant_x0020_Roles/Title','Relevant_x0020_Roles/Id','Department/Title','Department/Id')
                    .expand('Relevant_x0020_Roles','Department')
                    .filter(filterText)
                    .get()
                    .then((items: any[]) => {
                //return pnp.sp.web.getFolderByServerRelativeUrl(this._pageContext.web.serverRelativeUrl + '/SitePages').files.filter('').select('ServerRelativeUrl','Title','UniqueID','TimeCreated','Relevant_x0020_RolesId').get().then(items => {
                                        return Promise.resolve(items);
                }).catch(error => {
                    console.log('Error ensuring list: ' + error.message);
                    return Promise.reject('Error returning list');
                });
                /*
                return pnp.sp.web.lists.getByTitle(listName).getItemsByCAMLQuery(q).then((items: any[]) => {
                    return Promise.resolve(items);
                });
                */
        /*    }
        }).catch(error => {
            console.log('');
            return Promise.reject('error');
        });*/
    }

    public getRoles(): Promise<any[]> {
        return pnp.sp.web.lists.getByTitle('PolicyRoles').items.select('Title','ID').get().then((items: any[]) => {
            return Promise.resolve(items);
        }).catch(error => {
            console.log('Error ensuring list: ' + error.message);
            return Promise.reject('Error returning list');
        });
    }

    public getDepartments(): Promise<any[]> {
        return pnp.sp.web.lists.getByTitle('Departments').items.select('Title','ID').get().then((items: any[]) => {
            return Promise.resolve(items);
        }).catch(error => {
            console.log('Error ensuring list: ' + error.message);
            return Promise.reject('Error returning list');
        });
    }

    private _ensureList(listName: string): Promise<pnp.List> {
        if (listName) {
            // return pnp.sp.web.lists.ensure(listName).then((listEnsureResult) => Promise.resolve(listEnsureResult.list));
            return pnp.sp.web.lists.getByTitle(listName).get().then(list => {
                return Promise.resolve(list);
            }).catch(error => {
                return Promise.reject('Unable to find list ' + listName);
                // return Promise.resolve(null);
            });
        }
    }
}