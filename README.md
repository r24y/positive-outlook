# positive-outlook

Simple EWS client.

## Installation

```
npm i --save positive-outlook
```

## Usage

```js

import ExchangeClient, { Folder, Mailbox, Message } from 'positive-outlook';

const client = new ExchangeClient({
  username: 'my-email@outlook.com',
  password: 'correct horse battery staple',
  domain: 'ADMIN',
  strictSSL: false,  // in case you can't trust the certs
});

client.on('ready', () => {
  const inbox = Folder.Inbox();
  client::inbox.list().then(response => {
    response.messages.forEach(m => {
      console.log(m.isRead === 'true' ? '          ' : ' [unread] ', m.subject);
    });
  });

  const recipients = Mailbox.fromAddresses([
    'foo@example.com',
    '"Brian Bar" <bar@example.com>',
  ]);

  const message = new Message({
    to: recipients,
    subject: 'Hello world',
    body: 'Sending email via Node.js',
  });

  client::message.send().catch(err => console.error(err));
});
```

The API should be considered very unstable at this point. Once I've wrapped my head around most of the EWS API then you can start expecting some stability.

If you can't use the "bind" notation (`::`) then you can `call` the methods, e.g.:

```js
inbox.list.call(client).then(/* ... */);
message.send.call(client).catch(/* ... */);
```
