import Message from './EmailMessage';

const PRIVATE = Symbol();

export default class Folder {
  static Inbox() {
    return new Folder('inbox', { type: 'distinguished' });
  }
  static Sent() {
    return new Folder('sentitems', { type: 'distinguished' });
  }
  constructor(name, { type, id } = {}) {
    this[PRIVATE] = { name, type, id };
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
            messages: (messages || []).map(Message.fromResponse),
            raw: resp,
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
    return {
      FolderId: {
        attributes: this[PRIVATE].id,
      },
    };
  }

  static fromResponse(r) {
    const {
      FolderId: id,
      DisplayName: name,
      TotalCount: count,
      ChildFolderCount: childFolderCount,
      UnreadCount: unreadCount,
    } = r;
    const folder = new Folder(name, { id });
    Object.assign(folder[PRIVATE], { count, childFolderCount, unreadCount });
    return folder;
  }
}

const SPECIAL_FOLDERS = {
  Inbox: 'inbox',
  Sent: 'sentitems',
  Calendar: 'calendar',
  Outbox: 'outbox',
  Root: 'root',
  InstantMessageContacts: 'imcontactlist',
  Favorites: 'favorites',
  Junk: 'junkemail',
  Contacts: 'contacts',
  Drafts: 'drafts',
  Tasks: 'tasks',
  Trash: 'recoverableitemsroot',
  Archive: 'archiveroot',
  ArchiveInbox: 'archiveinbox',
}

Object.keys(SPECIAL_FOLDERS).forEach(k => {
  Folder[k] = function SpecialFolder() {
    return new Folder(SPECIAL_FOLDERS[k], { type: 'distinguished' });
  };
});
