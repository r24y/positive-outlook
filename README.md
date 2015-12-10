# positive-outlook

Simple EWS client.

## Installation

```
npm i --save positive-outlook
```

## Usage

```
import ExchangeClient from 'positive-outlook';
const client = new ExchangeClient({
  username: 'my-email@outlook.com',
  password: 'correct horse battery staple',
  domain: 'ADMIN',
  strictSSL: false,  // in case you can't trust the certs
});

client.on('ready', () => {
  client.listItems().then(response => {
    response.Items.Message.forEach(m => {
      console.log(m.IsRead === 'true' ? '          ' : ' [unread] ', m.Subject);
    });
  });
});
```

The API should be considered very unstable at this point. Once I've wrapped my head around most of the EWS API then you can start expecting some stability.
