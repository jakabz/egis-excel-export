import { sp } from "@pnp/sp";  
import { Web } from "@pnp/sp/webs";  
import "@pnp/sp/webs";
import "@pnp/sp/lists";  
import "@pnp/sp/items";  
import "@pnp/sp/site-users/web";
  
export class listService {

    public context:any;
  
    public setup(context: any): void {  
        sp.setup({  
            spfxContext: context  
        });
        this.context = context;
    }

    public async CurrentUserGroups(): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            try {
                sp.web.currentUser.groups().then((groups) => {
                    resolve(groups);
                });
            }
            catch (error) {
                console.log(error);
            }
        });
    }

    public async isGroupMember(spGroupsTitle:string[]):Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            try {
                if(spGroupsTitle.length === 0) {
                    resolve(true);
                } else {
                    let  isMember:boolean = false;
                    this.CurrentUserGroups().then(result => {
                        result.forEach(group => {
                            if(spGroupsTitle.indexOf(group.Title) > -1){
                                isMember = true;
                            }
                        });
                        resolve(isMember);
                    });
                }
            }
            catch (error) {
                console.log(error);
            }
        });
    }

    public async updateListItems(listId:string, items:number[]): Promise<any> { 
        return new Promise<any>(async (resolve, reject) => {  
            try {
                let list = sp.web.lists.getById(listId);
                const entityTypeFullName = await list.getListItemEntityTypeFullName();
                let batch = sp.web.createBatch();
                const now = (new Date()).toISOString();
                items.forEach(id => {
                    list.items.getById(id).inBatch(batch).update({ Completed: now, Status: 'Completed', CompletedByExport: 1 }, "*", entityTypeFullName).then(b => {
                        console.log(b);
                    });
                });
                batch.execute().then(result => {
                    resolve(result);
                }); 
            }  
            catch (error) {  
                console.log(error);  
            }  
        });  
    }

    public async createEmailListItems(items:any[]): Promise<any> {
        return new Promise<any>(async (resolve, reject) => {
            try {
                let list = sp.web.lists.getByTitle("EmailTechnicalList");
                const entityTypeFullName = await list.getListItemEntityTypeFullName();
                let batch = sp.web.createBatch();
                items.forEach(item => {
                    list.items.inBatch(batch).add({
                        Title: item.Title, 
                        ListName: item.ListName, 
                        Affiliate: item.Affiliate[0].lookupValue, 
                        RelatedItemID: item.RelatedItemID
                    }, entityTypeFullName).then(c => {
                        console.log(c);
                    });
                });
                batch.execute().then(result => {
                    resolve(result);
                }); 
            }
            catch (error) {
                console.log(error);
            }
        });
    }
    
}  
  
const SPListViewService = new listService();  
export default SPListViewService;