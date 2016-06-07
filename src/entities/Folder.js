import Message from './EmailMessage';

const PRIVATE = Symbol();

export default class Folder {
  static Inbox() {
    return new Folder('inbox', { type: 'distinguished' });
  }
  static Sent() {
    return new Folder('sentitems', { type: 'distinguished' });
  }
  constructor(name, { type } = {}) {
    this[PRIVATE] = {};
    this[PRIVATE].name = name;
    this[PRIVATE].type = type;
  }

  get list() {
    const folder = this;
    return function ({
      maxEntries = 30,
      offset = 0,
      basePoint = 'Beginning',
      allDetails = false,
    } = {}) {
      return new Promise((resolve, reject) => {
        this.client.FindItem({
          attributes: {Traversal: 'Shallow'},
          ItemShape: {
            BaseShape: allDetails ? 'AllProperties' : 'Default',
          },
          IndexedPageItemView: {
            attributes: {
              MaxEntriesReturned: maxEntries,
              BasePoint: basePoint,
              Offset: offset,
            },
          },
          ParentFolderIds: folder.asEws(),
        }, (err, resp) => {
          if (err) {
            const err2 = new Error(`Network error fetching folder '${folder[PRIVATE].name}'`);
            err2.original = err;
            return reject(err2);
          }
          const {
            ResponseMessages: {
              FindItemResponseMessage: {
                RootFolder: {
                  Items: items
                }
              }
            }
          } = resp;
          const {
            Message: messages,
          } = items;
          resolve({
            messages: messages.map(Message.fromResponse),
          });
        });

      });
    }
  }

  asEws() {
    if (this[PRIVATE].type === 'distinguished') {
      return {
        DistinguishedFolderId: {
          attributes: {
            Id: this[PRIVATE].name,
          },
        },
      };
    }
    throw new Error(`Unrecognized folder type ${this[PRIVATE].type}`);
  }
}
