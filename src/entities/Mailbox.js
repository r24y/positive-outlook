import { parseOneAddress } from 'email-addresses';

const PRIVATE = Symbol();

const XML_ELEMENT_NAMES = {
  name: 'Name',
  email: 'EmailAddress',
};

export default class Mailbox {
  constructor(name, { email } = {}) {
    this[PRIVATE] = { name, email };
  }

  asEws() {
    return { $xml: this.asXml() };
  }

  asXml() {
    const els = ['name', 'email'].map(k => [XML_ELEMENT_NAMES[k], this[PRIVATE][k]])
      .filter(([, v]) => v)
      .map(([k, v]) => `<t:${k}>${v}</t:${k}>`);
    return `<t:Mailbox>${els.join('')}</t:Mailbox>`;
  }

  static fromAddress(a) {
    const addr = parseOneAddress(a);
    return new Mailbox(a.name || a.local, { email: a.address });
  }

  static fromAddresses(addrs) {
    return addrs.map(Mailbox.fromAddress);
  }

  static fromResponse(r) {
    const mailbox = new Mailbox(r.Name, { email: r.EmailAddress });
    Object.assign(mailbox[PRIVATE], {
      routingType: r.RoutingType,
    });
    return mailbox;
  }
}
