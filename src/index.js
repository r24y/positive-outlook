import 'babel-polyfill';
import soap from 'soap-3';
import request from 'request';
import co from 'co';
import {EventEmitter} from 'events';
import _tmp from 'tmp';
import fs from 'fs-promise';
import esc from 'escape-html';

import os from 'os';
import path from 'path';

// Automatically delete any temporary files once we exit.
_tmp.setGracefulCleanup();

// Default to the server for outlook.com webmail.
const DEFAULT_EXCHANGE_ENDPOINT = 'https://outlook.office365.com/ews/Exchange.asmx';

// Handy references to the XML files to which we'll need to refer.
const SERVICES_FILE = path.join(__dirname, '..', 'wsdl', 'Services.wsdl');
const MESSAGES_FILE = path.join(__dirname, '..', 'wsdl', 'messages.xsd');
const TYPES_FILE = path.join(__dirname, '..', 'wsdl', 'types.xsd');

// tmp.dir() as a promise
function tmpdir(opt) {
  return new Promise((resolve, reject) => {
    _tmp.dir(opt, (err, res) => err ? reject(err) : resolve(res));
  });
}

const ns = {
  m: 'http://schemas.microsoft.com/exchange/services/2006/messages',
  t: 'http://schemas.microsoft.com/exchange/services/2006/types'
};

// Create a SOAP request as a promise
function mkSoap(wsdlPath, soapOptions) {
  return new Promise((resolve, reject) => {
    soap.createClient(wsdlPath, soapOptions, function (err, client, body) {
      err ? reject(err) : resolve(client);
    });
  });
}

// Utility function for a promise that resolves the next time `emitter` emits `event`.
function awaitEvent(emitter, event) {
  return new Promise((resolve) => {
    emitter.once(event, resolve);
  });
}

// The main workhorse.
class ExchangeClient extends EventEmitter {
  constructor({
    // The path to "Exchange.asmx" for the target server.
    exchangeEndpoint = DEFAULT_EXCHANGE_ENDPOINT,
    // Username as you would enter it for the webmail (omit domain).
    username,
    // Password.
    password,
    // Windows domain (e.g. "ADMIN").
    domain = '',
    // Hostname of your computer.
    workstation = '',
    // HTTP proxy
    proxy,
    // Set this to `false` for corporate intranets with self-signed certs.
    // TODO: allow user to specify a root CA instead of blindly accepting certs
    strictSSL = true,
  }) {
    super();

    // Unfortunately I don't think you can do fat-arrow generator functions. That would be so cool though!
    const that = this;

    // Use `co` for some async goodness. Thanks [tj](https://github.com/tj)!
    co(function *() {

      // Create a temp directory for our WSDL to live in. Unfortunately the [soap](https://www.npmjs.com/package/soap) module doesn't accept the WSDL as a string, so we have to go through this tempfile song and dance. No big deal.
      const tmpPath = yield tmpdir();
      const wsdlPath = path.join(tmpPath, 'Services.wsdl');

      // Read in the content for our XML files.
      const [wsdl, messages, types] = yield [
        fs.readFile(SERVICES_FILE),
        fs.readFile(MESSAGES_FILE),
        fs.readFile(TYPES_FILE),
      ];

      // Look for the placeholder in `Services.wsdl` and replace it with our desired Exchange endpoint.
      const wsdl2 = wsdl.toString().replace('%%EXCHANGE_ASMX_ENDPOINT_LOCATION%%', exchangeEndpoint);
      yield [
        fs.writeFile(wsdlPath, wsdl2),
        fs.writeFile(path.join(tmpPath, 'messages.xsd'), messages),
        fs.writeFile(path.join(tmpPath, 'types.xsd'), types),
      ];

      const soapOptions = {
        wsdl_options: {
          ntlm: true,
          username,
          password,
          domain,
          workstation,
          strictSSL,
          proxy,
          agentOptions: {
            rejectUnauthorized: strictSSL,
          },
          rejectUnauthorized: strictSSL,
        },
        request: request.defaults({
          strictSSL: false,
        })
      };

      const client = that.client = yield mkSoap(wsdlPath, soapOptions);
      client.setSecurity(new soap.NtlmSecurity(soapOptions.wsdl_options));

      // ok we're all set!
      that.emit('ready');
    }).catch(err => console.error(err.stack));
  }

  // ## listItems
  // List items from the inbox.
  listItems({
    // Number of items to take.
    maxEntries = 30,
    // Number of items to skip (useful for paging).
    offset = 0,
    // EWS property I decided to expose here. Not really sure what it does.
    basePoint = 'Beginning',
    // Set to `true` to pull all available info, `false` to pull some basic info.
    allDetails = false,
    // The folder to read. I don't think this is set up properly (may need to swap out `DistinguishedFolderId`) but here you go.
    folder = 'inbox',
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
        ParentFolderIds: {
          DistinguishedFolderId: {
            attributes: {Id: folder},
          },
        },
      }, (err, resp) => {
        if (err) {
          const err2 = new Error(`Network error fetching folder '${folder}'`);
          err2.original = err;
          return reject(err2);
        }
        resolve(resp.ResponseMessages.FindItemResponseMessage.RootFolder);
      });
    });
  }

  createMessage({
    recipients = [],
    subject = 'New Message',
    body = '',
  } = {}) {
    return new Promise((resolve, reject) => {
      this.client.CreateItem({
        attributes: { MessageDisposition: 'SendAndSaveCopy' },
        SavedItemFolderId: { DistinguishedFolderId: { attributes: { Id: 'sentitems' } } },
        Items: {
          $xml: `
          <t:Message>
            <t:Subject>${subject}</t:Subject>
            <t:Body BodyType="HTML">${esc(body)}</t:Body>
            <t:ToRecipients>
            ${
              recipients.map(r => `<t:Mailbox><t:EmailAddress>${r}</t:EmailAddress></t:Mailbox>`)
                .join()
            }
            </t:ToRecipients>
          </t:Message>
          `
        }
      }, (err, resp) => {
        if (err) {
          const err2 = new Error(`Network error sending message`);
          err2.original = err;
          return reject(err2);
        }
        resolve(resp);
      });
    });
  }
}

module.exports = ExchangeClient;
module.exports.awaitEvent = awaitEvent;

// If we've directly called this file, fetch the user's inbox.
// This reads the values of `EXCHANGE_ENDPT`, `OUTLOOK_USER`, `OUTLOOK_PASS`, `OUTLOOK_DOMAIN` from
// your environment and uses them to log in.
if (require.main === module) {
  if (!(process.env.EXCHANGE_ENDPT && process.env.OUTLOOK_USER && process.env.OUTLOOK_PASS && process.env.OUTLOOK_DOMAIN)) {
    console.log('Need to set environment variables EXCHANGE_ENDPT, OUTLOOK_USER, OUTLOOK_PASS, OUTLOOK_DOMAIN');
    process.exit(0);
  }
  co(function *() {
    const ews = new ExchangeClient({
      exchangeEndpoint: process.env.EXCHANGE_ENDPT,
      username: process.env.OUTLOOK_USER,
      password: process.env.OUTLOOK_PASS,
      domain: process.env.OUTLOOK_DOMAIN,
      // Assume corporate intranet.
      strictSSL: false,
    });

    // Use our utility function to wait for the client to be ready.
    yield awaitEvent(ews, 'ready');

    // Perform some simple formatting and spew the top 30 messages to the console.
    console.log((yield ews.listItems()).Items.Message.map(m => `${m.IsRead==='true' ? '          ' : ' [unread] '} ${m.Subject}`).join('\n'));

  }).catch(err => console.error(err.original.stack));
}
