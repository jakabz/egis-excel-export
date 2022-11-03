export class urlQueryParameters {

    public getValue = (name: string) => {
        const search = document.location.search.slice(1,-1).split('&');
        let result:string;
        search.forEach(item => {
            const key = item.split('=')[0];
            const value = item.split('=')[1];
            if(key===name){
                result = value; 
            }
        });
        return result;
    }

}

const UrlQueryParameters = new urlQueryParameters();  
export default UrlQueryParameters;