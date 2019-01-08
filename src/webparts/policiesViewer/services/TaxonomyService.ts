import { taxonomy, ITermStore, ITermSet, ITerms } from '@pnp/sp-taxonomy';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as pnp from 'sp-pnp-js';

export interface ITaxonomyService {
    _pageContext: PageContext;
    getDepartments: () => Promise<ITerms>;
}

// No longer used
export class TaxonomyService implements ITaxonomyService {
    public static readonly serviceKey: ServiceKey<ITaxonomyService> = ServiceKey.create<ITaxonomyService>('ImageGallery:TaxonomyService', TaxonomyService);
    public _pageContext: PageContext;    

    constructor(serviceScope: ServiceScope) {
        

        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
        });
    }

    public async getDepartments(): Promise<ITerms> {
        const stores = await taxonomy.termStores.get(); //.select("Name")
        const storeName = stores[0].Name;
        const store: ITermStore = taxonomy.termStores.getByName(storeName);
        const termsets = await store.getTermSetsByName("PnP-Organizations", 1033);
        const termsetList = await termsets.get();
        return termsets.getByName('PnP-Organizations').get().then(termset => {
            return Promise.resolve(termset.terms);
        }).catch(error => {
            return Promise.reject('Error retrieving termsets: ' + error.message);
        });
        //return setWithData;
    }
}