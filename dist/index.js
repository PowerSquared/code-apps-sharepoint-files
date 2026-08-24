import { getClient } from '@microsoft/power-apps/data';
import { FILE_APIS } from './operations';
export { FILE_APIS } from './operations';
function unwrap(result, action) {
    if (!result.success) {
        throw new Error(`${action} failed: ${result.error?.message ?? 'unknown error'}`);
    }
    return result.data;
}
/**
 * Reads a blob as base64. A `format: 'binary'` body parameter must be a base64 string, which the
 * client library decodes to bytes and sends as application/octet-stream.
 *
 * Uses FileReader, so this package is browser-only — which every code app is.
 */
function toBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(reader.error ?? new Error('Could not read file.'));
        reader.onload = () => {
            const dataUrl = reader.result;
            const comma = dataUrl.indexOf(',');
            resolve(comma === -1 ? dataUrl : dataUrl.slice(comma + 1));
        };
        reader.readAsDataURL(blob);
    });
}
function resolveFileName(file, options) {
    const name = options.fileName ?? file.name;
    if (!name) {
        throw new Error('A file name is required. Pass options.fileName when uploading a Blob.');
    }
    return name;
}
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
export function createSharePointFileService(options) {
    const { dataSourcesInfo, dataSourceName } = options;
    const registry = dataSourcesInfo;
    const dataSource = registry[dataSourceName];
    if (!dataSource) {
        throw new Error(`Data source '${dataSourceName}' is not in dataSourcesInfo. ` +
            `Available: ${Object.keys(registry).join(', ') || '(none)'}.`);
    }
    // Register onto the *same* object the generated services import. The library's
    // PowerDataSourcesInfoProvider keeps the first dataSourcesInfo handed to it and ignores every
    // later one, so mutating in place makes registration independent of which service issues the
    // first request.
    Object.assign(dataSource.apis, FILE_APIS);
    const client = getClient(dataSourcesInfo);
    return {
        async uploadFile(file, uploadOptions) {
            const result = await client.executeAsync({
                connectorOperation: {
                    tableName: dataSourceName,
                    operationName: 'CreateFile',
                    parameters: {
                        folderPath: uploadOptions.folderPath,
                        name: resolveFileName(file, uploadOptions),
                        body: await toBase64(file),
                    },
                },
            });
            return unwrap(result, 'Upload');
        },
        async replaceFileContent(fileId, file) {
            const result = await client.executeAsync({
                connectorOperation: {
                    tableName: dataSourceName,
                    operationName: 'UpdateFile',
                    parameters: { id: fileId, body: await toBase64(file) },
                },
            });
            return unwrap(result, 'Replace');
        },
        async deleteFile(fileId) {
            const result = await client.executeAsync({
                connectorOperation: {
                    tableName: dataSourceName,
                    operationName: 'DeleteFile',
                    parameters: { id: fileId },
                },
            });
            unwrap(result, 'Delete');
        },
        async getFileMetadata(fileId) {
            const result = await client.executeAsync({
                connectorOperation: {
                    tableName: dataSourceName,
                    operationName: 'GetFileMetadata',
                    parameters: { id: fileId },
                },
            });
            return unwrap(result, 'Get metadata');
        },
        async getFileContent(fileId) {
            const result = await client.executeAsync({
                connectorOperation: {
                    tableName: dataSourceName,
                    operationName: 'GetFileContent',
                    parameters: { id: fileId, inferContentType: true },
                },
            });
            return unwrap(result, 'Download');
        },
    };
}
//# sourceMappingURL=index.js.map