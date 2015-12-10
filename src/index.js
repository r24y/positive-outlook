import soap from '@r24y/soap';
import request from 'request';
import co from 'co';
import {EventEmitter} from 'events';
import _tmp from 'tmp';
import fs from 'fs-promise';

import path from 'path';

_tmp.setGracefulCleanup();

const DEFAULT_EXCHANGE_ENDPOINT = 'https://outlook.office365.com/ews/Exchange.asmx';
const SERVICES_FILE = path.join(__dirname, '..', 'wsdl', 'Services.wsdl');
const MESSAGES_FILE = path.join(__dirname, '..', 'wsdl', 'messages.xsd');
const TYPES_FILE = path.join(__dirname, '..', 'wsdl', 'types.xsd');

function tmpdir(opt) {
  return new Promise((resolve, reject) => {
    _tmp.dir(opt, (err, res) => err ? reject(err) : resolve(res));
  });
}

function mkSoap(wsdlPath, soapOptions) {
  return new Promise((resolve, reject) => {
    soap.createClient(wsdlPath, soapOptions, function (err, client, body) {
      err ? reject(err) : resolve(client);
    });
  });
}

function awaitEvent(emitter, event) {
  return new Promise((resolve) => {
    emitter.once(event, resolve);
  });
}

class ExchangeClient extends EventEmitter {
  constructor({
    exchangeEndpoint = DEFAULT_EXCHANGE_ENDPOINT,
    username,
    password,
    domain = '',
    workstation = '',
    proxy,
    strictSSL = true,
  }) {
    super();
    const that = this;
    co(function *() {
      const tmpPath = yield tmpdir();
      const wsdlPath = path.join(tmpPath, 'Services.wsdl');
      const [wsdl, messages, types] = yield [
        fs.readFile(SERVICES_FILE),
        fs.readFile(MESSAGES_FILE),
        fs.readFile(TYPES_FILE),
      ];
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
      that.emit('ready');
    }).catch(err => console.error(err.stack));
  }
  listItems({
    maxEntries = 30,
    offset = 0,
    basePoint = 'Beginning',
    allDetails = false,
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
}

co(function *() {
  const ews = new ExchangeClient({
    exchangeEndpoint: process.env.EXCHANGE_ENDPT,
    username: process.env.USER,
    password: process.env.PASS,
    domain: process.env.DOMAIN,
    strictSSL: false,
  });
  yield awaitEvent(ews, 'ready');
  console.log((yield ews.listItems()).Items.Message.map(m => `${m.IsRead==='true' ? '          ' : ' [unread] '} ${m.Subject}`).join('\n'));

}).catch(err => console.error(err.original.stack));
