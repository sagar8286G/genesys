// import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from './pnpjs.preset';
// import { sp, } from '@pnp/sp';
import { Web } from '@pnp/sp/webs';
// import { graph, } from "@pnp/graph";
// import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, HttpClient, MSGraphClient } from '@microsoft/sp-http';
// import * as $ from 'jquery';
import { registerDefaultFontFaces } from "@uifabric/styling";
import * as moment from 'moment';
// import { SiteUser } from "@pnp/sp/src/siteusers";
// import { dateAdd } from "@pnp/common";
import { escape, update } from '@microsoft/sp-lodash-subset';

export interface IAllItems {
    listName: string;
    Id?: string;
    selectQuery?: string;
    filterQuery?: string;
    expandQuery?: string;
    orderByQuery?: { columnName: string, ascending: boolean };
    topQuery?: number;
}

// Class Services
export default class spservices {

    // private graphClient: MSGraphClient = null;

    constructor(private context: any) {
        // Setuo Context to PnPjs and MSGraph
        sp.setup({
            spfxContext: this.context
        });

        // graph.setup({
        //     spfxContext: this.context
        // });
        // Init
        this.onInit();
    }
    // OnInit Function
    private async onInit() {
    }



    public async getAllListItems(Item: IAllItems): Promise<any[]> {
        try {
            const orderByColumn = Item.orderByQuery ? Item.orderByQuery.columnName : 'Id';
            const orderByAscending = Item.orderByQuery ? Item.orderByQuery.ascending : true;
            return await sp.web.lists.getByTitle(Item.listName).items
                .filter(Item.filterQuery ? Item.filterQuery : '')
                .select(Item.selectQuery ? Item.selectQuery : '*')
                .expand(Item.expandQuery ? Item.expandQuery : '')
                .top(Item.topQuery ? Item.topQuery : 0)
                .orderBy(orderByColumn, orderByAscending).get();
        } catch (error) {
            return Promise.reject(error);
        }
    }


}
