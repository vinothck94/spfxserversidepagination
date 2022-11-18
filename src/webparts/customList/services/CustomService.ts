import { Web, IWeb } from "@pnp/sp/webs";
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions
} from '@microsoft/sp-http';




interface ICustomProperty {
  context: any;
  currentPage: number;
  recordsPerPage: number;
  properties?: string;
  filter?: string;
  expand?: string;
}



export default class CustomService {
  public webUrl: string = '';
  public listName: string = '';

  public getList(customProperty: ICustomProperty, callBack: any) {
    var take = customProperty.recordsPerPage;
    var skip = (customProperty.currentPage * customProperty.recordsPerPage) - customProperty.recordsPerPage;
    var web: IWeb = Web(this.webUrl);

    ////Get count of the list
    var getUrl = customProperty.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('" + this.listName + "')/ItemCount";

    customProperty.context.spHttpClient.get(getUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            if (responseJSON != null && responseJSON.value != null) {
              let itemCount: number = parseInt(responseJSON.value.toString());

              ////Get List data with pagination
              if (!customProperty.properties) {
                customProperty.properties = '';
              }
              if (!customProperty.filter) {
                customProperty.filter = '';
              }
              if (!customProperty.expand) {
                customProperty.expand = '';
              }

              web.lists
                .getByTitle(this.listName)
                .items.select(customProperty.properties)
                .expand(customProperty.expand)
                .filter(customProperty.filter)
                .skip(skip)
                .top(take)
                .get().then(res => {
                  callBack({
                    count: itemCount,
                    data: res
                  });
                });

            }
          });
        }
      });
  }


}
