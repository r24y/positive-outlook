import co from 'co';

import ExchangeClient, {awaitEvent} from '../src';
import Folder from '../src/entities/Folder';

console.log('Starting up');

// This reads the values of `EXCHANGE_ENDPT`, `OUTLOOK_USER`, `OUTLOOK_PASS`, `OUTLOOK_DOMAIN` from
// your environment and uses them to log in.
if (!(process.env.EXCHANGE_ENDPT && process.env.OUTLOOK_USER && process.env.OUTLOOK_PASS && process.env.OUTLOOK_DOMAIN)) {
  console.log('Need to set environment variables EXCHANGE_ENDPT, OUTLOOK_USER, OUTLOOK_PASS, OUTLOOK_DOMAIN');
  process.exit(0);
}

co(function *() {
  console.log('Creating client');
  const ews = new ExchangeClient({
    exchangeEndpoint: process.env.EXCHANGE_ENDPT,
    username: process.env.OUTLOOK_USER,
    password: process.env.OUTLOOK_PASS,
    domain: process.env.OUTLOOK_DOMAIN,
    // Assume corporate intranet.
    strictSSL: false,
  });
  console.log('Waiting for connection');

  // Use our utility function to wait for the client to be ready.
  yield awaitEvent(ews, 'ready');

/*
  console.log('Creating message');

  try {
    console.log(yield ews.createMessage({
      recipients: ['me@mysite.com'],
      subject: 'Hello styled NodeJS email!',
      body: `
      <style>body { font-family: 'Fira Sans', 'Helvetica', sans-serif; }</style>
      Welcome to <em>the future</em> of <strong>productivity</strong>
      `,
    }));
  } catch (err) {
    console.error(err);
    console.error(err.stack);
  }
  */

  console.log('Listing inbox');

  const inbox = Folder.Inbox();

  // Perform some simple formatting and spew the top 30 messages to the console.
  console.log((yield ews::inbox.list()).messages.join('\n'));

  //console.log(ews.soap);

}).catch(err => console.error(err.original.stack));
