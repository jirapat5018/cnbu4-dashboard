/** Prototype connector: reads the Excel from the dashboard's own web server (./data/...). Same contract as SharePointConnector. */
import { ConnectorError } from './sharepointConnector.js';

export class SampleConnector {
  constructor(cfg, fetchImpl = (...a) => fetch(...a)) { this.cfg = cfg; this.fetch = fetchImpl; }

  _meta(res) {
    return {
      version: res.headers.get('etag') || res.headers.get('last-modified') || String(Date.now()),
      lastModified: res.headers.get('last-modified') ? new Date(res.headers.get('last-modified')).toISOString() : null,
      name: this.cfg.url.split('/').pop(),
    };
  }

  async getMeta() {
    let res;
    try { res = await this.fetch(this.cfg.url, { method: 'HEAD', cache: 'no-store' }); }
    catch (e) { throw new ConnectorError('เปิดไฟล์ตัวอย่างไม่ได้', { code: 'network', cause: e }); }
    if (!res.ok) throw new ConnectorError(`ไม่พบไฟล์ตัวอย่าง (${res.status}) ${this.cfg.url}`, { status: res.status });
    return this._meta(res);
  }

  async fetchFile() {
    let res;
    try { res = await this.fetch(this.cfg.url, { cache: 'no-store' }); }
    catch (e) { throw new ConnectorError('เปิดไฟล์ตัวอย่างไม่ได้', { code: 'network', cause: e }); }
    if (!res.ok) throw new ConnectorError(`ไม่พบไฟล์ตัวอย่าง (${res.status}) ${this.cfg.url}`, { status: res.status });
    return { buffer: await res.arrayBuffer(), meta: this._meta(res) };
  }
}
