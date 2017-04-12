import { ISPListItem } from './SpOperation';

export default class MockHttpClient {

    private static _items: ISPListItem[] = [
        { Title: 'Document 1', Id: '1', ServerRelativeUrl :'', Name:'Document'},
        { Title: 'Document 2', Id: '2', ServerRelativeUrl :'', Name:'Document'},
        { Title: 'Document 3', Id: '3', ServerRelativeUrl :'', Name:'Document'}
    ];

    public static get(): Promise<ISPListItem[]> {
        return new Promise<ISPListItem[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}