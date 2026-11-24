/** Connector for "no server" use: the user picks the Excel from disk (works when index.html is opened by double-click). */
import { ConnectorError } from './sharepointConnector.js';

export class LocalFileConnector {
  constructor(input) { this.input = input; }

  /** must be called from a click handler (browser rule) */
  pick() {
    return new Promise((resolve) => {
      this.input.value = '';
      this.input.onchange = () => resolve(this.input.files[0] || null);
      this.input.click();
    });
  }

  async getMeta() { return { version: String(this.file?.lastModified || ''), lastModified: this.file ? new Date(this.file.lastModified).toISOString() : null, name: this.file?.name }; }

  async fetchFile() {
    this.file = await this.pick();
    if (!this.file) throw new ConnectorError('ยังไม่ได้เลือกไฟล์ Excel', { code: 'cancel' });
    if (!/\.xls[xm]$/i.test(this.file.name)) throw new ConnectorError('ไฟล์ต้องเป็น .xlsx', { code: 'type' });
    return { buffer: await this.file.arrayBuffer(), meta: await this.getMeta() };
  }
}
