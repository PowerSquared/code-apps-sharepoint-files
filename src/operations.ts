import type { ApiDefinition, ApiParameter } from './types';

/*
 * SharePoint connector operations the CLI does not generate.
 *
 * `connectionId`, `dataset` and `tableName` are synthetic: the executor fills them from the
 * connection reference and strips them before matching caller-supplied parameters. They are
 * declared so the path placeholders resolve, but must never be passed by callers. This is why
 * these operations automatically inherit a dataset bound to an `@envvar:` environment variable.
 *
 * Operation IDs and parameter keys come from the SharePoint connector reference:
 * https://learn.microsoft.com/en-us/connectors/sharepointonline/
 */

const CONNECTION: ApiParameter = { name: 'connectionId', in: 'path', required: true, type: 'string' };
const DATASET: ApiParameter = { name: 'dataset', in: 'path', required: true, type: 'string' };
const FILE_ID: ApiParameter = { name: 'id', in: 'path', required: true, type: 'string' };

/**
 * Responses are only JSON-parsed when the operation declares `responseInfo` for the status code;
 * without it the raw response string is returned. `application/octet-stream` responses are always
 * returned as a `Uint8Array` regardless, which is why GetFileContent omits this.
 */
const JSON_OBJECT_RESPONSE = { '200': { type: 'object' } };

export const FILE_APIS: Record<string, ApiDefinition> = {
  CreateFile: {
    path: '/{connectionId}/datasets/{dataset}/files',
    method: 'POST',
    parameters: [
      CONNECTION,
      DATASET,
      { name: 'folderPath', in: 'query', required: true, type: 'string' },
      { name: 'name', in: 'query', required: true, type: 'string' },
      // A `format: 'binary'` body whose value is a string is base64-decoded to bytes and sent as
      // application/octet-stream. Anything else is JSON.stringify'd and corrupts the upload.
      { name: 'body', in: 'body', required: true, type: 'string', format: 'binary' },
    ],
    responseInfo: JSON_OBJECT_RESPONSE,
  },
  UpdateFile: {
    path: '/{connectionId}/datasets/{dataset}/files/{id}',
    method: 'PUT',
    parameters: [
      CONNECTION,
      DATASET,
      FILE_ID,
      { name: 'body', in: 'body', required: true, type: 'string', format: 'binary' },
    ],
    responseInfo: JSON_OBJECT_RESPONSE,
  },
  DeleteFile: {
    path: '/{connectionId}/datasets/{dataset}/files/{id}',
    method: 'DELETE',
    parameters: [CONNECTION, DATASET, FILE_ID],
  },
  GetFileMetadata: {
    path: '/{connectionId}/datasets/{dataset}/files/{id}',
    method: 'GET',
    parameters: [CONNECTION, DATASET, FILE_ID],
    responseInfo: JSON_OBJECT_RESPONSE,
  },
  GetFileContent: {
    path: '/{connectionId}/datasets/{dataset}/files/{id}/content',
    method: 'GET',
    parameters: [
      CONNECTION,
      DATASET,
      FILE_ID,
      { name: 'inferContentType', in: 'query', required: false, type: 'boolean' },
    ],
  },
};
