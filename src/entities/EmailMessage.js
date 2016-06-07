import esc from 'escape-html';
import Item from './Item';
import Mailbox from './Mailbox';

const PRIVATE = Symbol();

export default class EmailMessage extends Item {
  constructor({
    to = [],
    cc = [],
    bcc = [],
    subject = '<no subject>',
    body = '',
    format = 'HTML',
  } = {}) {
    super();
    this[PRIVATE] = {
      to, cc, bcc, subject, body, format,
    };
  }

  get subject() {
    return this[PRIVATE].subject;
  }

  get from() {
    return this[PRIVATE].from;
  }

  get isRead() {
    return this[PRIVATE].isRead;
  }

  get body() {
    return this[PRIVATE].body;
  }

  get send() {
    const message = this;
    function send() {
      return new Promise((resolve, reject) => {
        this.client.CreateItem({
          attributes: { MessageDisposition: 'SendAndSaveCopy' },
          SavedItemFolderId: { DistinguishedFolderId: { attributes: { Id: 'sentitems' } } },
          Items: message.asEws(),
        }, (err, resp) => {
          if (err) {
            const err2 = new Error(`Network error sending message`);
            err2.original = err;
            err2.message = message;
            return reject(err2);
          }
          resolve(resp);
        });
      });
    }

    return send;
  }

  get fetch() {
    const message = this;
    function fetch() {
      return new Promise((resolve, reject) => {
        this.client.GetItem({
          ItemShape: {
            BaseShape: 'Default',
            IncludeMimeContent: true,
          },
          ItemIds: {
            ItemId: message[PRIVATE].id,
          }
        }, (err, resp) => {
          if (err) {
            const err2 = new Error(`Network error fetching message`);
            err2.original = err2;
            err2.message = message;
            return reject(err2);
          }
          const {
            ResponseMessages: {
              GetItemResponseMessage: {
                Items: {
                  Message: messages
                }
              },
            },
          } = resp;
          const message = messages.length ? messages[0] : messages;
          resolve(EmailMessage.fromResponse(message));
        });
      })
    }
    return fetch;
  }

  toString() {
    return `${this[PRIVATE].isRead === false ? '[*]' : '   '} ${this[PRIVATE].subject}`;
  }

  asEws() {
    const { to, cc, bcc, subject, body, format } = this[PRIVATE];
    return {
      $xml: `
        <t:Message>
          <t:Subject>${subject}</t:Subject>
          <t:Body BodyType="${format}">${esc(body)}</t:Body>
          <t:ToRecipients>${to.map(r => r.asXml()).join('')}</t:ToRecipients>
        </t:Message>
      `,
    };
  }

  static fromResponse(m) {
    const message = new EmailMessage({
      subject: m.Subject,
      to:
        typeof m.ToRecipients.Mailbox.map === 'function'
          ? m.ToRecipients.Mailbox.map(Mailbox.fromResponse)
          : Mailbox.fromResponse(m.ToRecipients.Mailbox),
    });
    const {
      ItemId: id,
      Importance: importance = 'Normal',
      Sensitivity: sensitivity,
      Body: {
        attributes: {
          BodyType = 'text',
        } = {},
        $value: Body$Value,
      } = {},
    } = m;
    Object.assign(message[PRIVATE], {
      // available in brief form
      isRead: m.IsRead === 'true',
      size: Number(m.Size),
      sensitivity,
      from: Mailbox.fromResponse(m.From.Mailbox),
      id,
      tsSent: new Date(m.DateTimeSent),
      tsCreated: new Date(m.DateTimeCreated),
      hasAttachments: m.HasAttachments === 'true',

      // available only if you've fetched the message
      body: Body$Value || m.Body,
      format: m.Body && BodyType,
      importance: m.Importance && importance,
    });
    return message;
  }
}

/*
{
        "ItemId": {
          "attributes": {
            "Id": "AAAYAHJ5YW4ubXVsbGVyQG5vdmFydGlzLmNvbQBGAAAAAACo5Av9tSeVT6YAZIyHb2y8BwDvJKw/sR1pR5VkjXrl3IYPAAAASeqMAABsL+MJivdrR6ZWk4hmeFgSAAGIvJIVAAA=",
            "ChangeKey": "CQAAABYAAABsL+MJivdrR6ZWk4hmeFgSAAGIzxow"
          }
        },
        "Subject": "Foo Bar",
        "Sensitivity": "Normal",
        "Size": "74162",
        "DateTimeSent": "2016-06-07T16:04:27Z",
        "DateTimeCreated": "2016-06-07T16:04:29Z",
        "HasAttachments": "false",
        "From": {
          "Mailbox": {
            "Name": "Person McPersonFace",
            "EmailAddress": "person@person.com",
            "RoutingType": "SMTP"
          }
        },
        "IsRead": "false"
      },
      */
