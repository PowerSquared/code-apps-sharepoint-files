import type { CreateSharePointFileServiceOptions, SharePointFileService } from './types';
export type { ApiDefinition, ApiParameter, CreateSharePointFileServiceOptions, GeneratedDataSourcesInfo, SharePointFile, SharePointFileService, UploadOptions, } from './types';
export { FILE_APIS } from './operations';
/**
 * Registers the SharePoint file operations against a code app's data source and returns a service
 * for calling them.
 *
 * Requires `@microsoft/power-apps` >= 1.3.0: earlier versions have no binary-body handling, so
 * uploads are JSON-encoded and the resulting files are corrupt.
 *
 * @example
 * ```ts
 * import { createSharePointFileService } from 'code-apps-sharepoint-files';
 * import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
 *
 * const files = createSharePointFileService({ dataSourcesInfo, dataSourceName: 'mylibrary' });
 * await files.uploadFile(file, { folderPath: '/My Library' });
 * ```
 */
export declare function createSharePointFileService(options: CreateSharePointFileServiceOptions): SharePointFileService;
//# sourceMappingURL=index.d.ts.map