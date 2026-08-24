/** Mirrors the client library's internal IApiDefinition shape, which is not exported publicly. */
export interface ApiParameter {
    name: string;
    in: 'path' | 'query' | 'body' | 'header';
    required: boolean;
    type: string;
    format?: string;
}
export interface ApiDefinition {
    path: string;
    method: string;
    parameters: ApiParameter[];
    responseInfo?: Record<string, {
        type?: string;
        format?: string;
    }>;
}
/**
 * Structural type for the generated `dataSourcesInfo` object a code app imports from
 * `.power/schemas/appschemas/dataSourcesInfo`. Declared structurally so this package never has to
 * reach into a consuming project's generated files.
 */
export type GeneratedDataSourcesInfo = Record<string, {
    tableId: string;
    apis: Record<string, unknown>;
}>;
/** Subset of the connector's SPBlobMetadataResponse returned by the file operations. */
export interface SharePointFile {
    Id: string;
    Name: string;
    DisplayName?: string;
    Path?: string;
    LastModified?: string;
    Size?: number;
    MediaType?: string;
    ETag?: string;
    IsFolder?: boolean;
}
export interface UploadOptions {
    /**
     * Destination folder, as a path starting with an existing library — for example
     * `/Property Listing Photos`. The connector's file operations address folders by path, while its
     * table operations address the library by GUID, so a library ID cannot be used here.
     */
    folderPath: string;
    /** File name to create. Defaults to `file.name` when a `File` is supplied. */
    fileName?: string;
}
export interface SharePointFileService {
    /** Uploads a file and returns its SharePoint metadata. */
    uploadFile(file: Blob, options: UploadOptions): Promise<SharePointFile>;
    /** Replaces an existing file's contents, keeping its identifier and list item. */
    replaceFileContent(fileId: string, file: Blob): Promise<SharePointFile>;
    /** Permanently deletes a file. */
    deleteFile(fileId: string): Promise<void>;
    /** Reads file metadata (size, etag, path). */
    getFileMetadata(fileId: string): Promise<SharePointFile>;
    /** Downloads file bytes. */
    getFileContent(fileId: string): Promise<Uint8Array>;
}
export interface CreateSharePointFileServiceOptions {
    /**
     * The project's generated `dataSourcesInfo`. Pass the object itself, not a copy: the client
     * library's data source registry is a first-call-wins singleton, so operations are registered by
     * mutating the same object the generated services import.
     */
    dataSourcesInfo: GeneratedDataSourcesInfo;
    /** Data source name of the SharePoint library, e.g. `propertylistingphotos`. */
    dataSourceName: string;
}
//# sourceMappingURL=types.d.ts.map