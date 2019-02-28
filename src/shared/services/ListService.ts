import { Text }                                                 	from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import { sp,Web } from '@pnp/sp';
export default class ListService {
 private web:Web;
    constructor(webUrl:string) {
		this.web=new Web(webUrl);
    }

	public getListItemsByQuery(listId: string, listFields: string[], expandFields: string[], top: number=10): Promise<any> {
		return new Promise<any>((resolve, reject) => {
			if (expandFields!==null) {
				this.web.lists.getByTitle(listId).items.select(...listFields)
				.expand(...expandFields).top(top).get().then(response => resolve(response))
				.catch(error => reject(error));
			} else {
				this.web.lists.getByTitle(listId).items.select(...listFields)
				.top(top).get().then(response => resolve(response))
				.catch(error => reject(error));

			}

		});
	}

}


export interface IListTitle {
	id: string;
	title: string;
}