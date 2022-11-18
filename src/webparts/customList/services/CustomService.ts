import { Web, IWeb } from "@pnp/sp/webs";
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions
} from '@microsoft/sp-http';

export default class CustomService {
  public webUrl: string = '';
  public listName: string = '';

  public async getlist(context: WebPartContext, currentPage: number, recordsPerPage: number, properties: string, callBack: any) {

    var take = recordsPerPage;
    var skip = (currentPage * recordsPerPage) - recordsPerPage;
    var web: IWeb = Web(this.webUrl);

    ////Get count of the list
    var getUrl = context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('" + this.listName + "')/ItemCount";

    context.spHttpClient.get(getUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            if (responseJSON != null && responseJSON.value != null) {
              let itemCount: number = parseInt(responseJSON.value.toString());

              ////Get List data with pagination
              web.lists
                .getByTitle(this.listName)
                .items.select()
                .skip(skip)
                .top(take)
                .get().then(res => {
                  callBack({
                    count: itemCount,
                    data: res
                  });
                })
            }
          });
        }
      });
  }
}
