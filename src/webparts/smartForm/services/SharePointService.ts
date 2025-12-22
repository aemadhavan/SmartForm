import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISharePointUser {
  displayName: string;
  email: string;
  loginName: string;
}

export interface ISharePointFileInfo {
  name: string;
  serverRelativeUrl: string;
  timeLastModified: string;
  length?: number;
}

export interface ISharePointFileInfoPage<T extends ISharePointFileInfo = ISharePointFileInfo> {
  items: T[];
  nextLink?: string;
}

export interface IEnsureTextFieldRequest {
  internalName: string;
  displayName: string;
  maxLength?: number;
}

export class SharePointService {
  private context: WebPartContext;
  private _libraryFieldsCache: { [libraryName: string]: { [internalName: string]: true } } = {};

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public static toInternalName(fieldName: string): string {
    // SharePoint internal names should be alphanumeric/underscore.
    // Keep this deterministic so the same form field maps to the same column.
    const cleaned = (fieldName || '')
      .trim()
      .replace(/[^A-Za-z0-9_]/g, '_')
      .replace(/_+/g, '_');

    // Ensure it doesn't start with a number
    const safe = cleaned.match(/^[0-9]/) ? `F_${cleaned}` : cleaned;
    return safe || 'Field';
  }

  private escapeODataValue(value: string): string {
    return value.replace(/'/g, "''");
  }

  private async getLibraryRootFolderServerRelativeUrl(libraryName: string): Promise<string> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const safeLibraryName = this.escapeODataValue(libraryName);

    const listEndpoint = `${webUrl}/_api/web/lists/getbytitle('${safeLibraryName}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;

    const listResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
      listEndpoint,
      SPHttpClient.configurations.v1
    );

    if (!listResponse.ok) {
      throw new Error(
        `Library '${libraryName}' not found or access denied. Please check the library name and permissions.`
      );
    }

    const listData = await listResponse.json();
    return listData?.RootFolder?.ServerRelativeUrl as string;
  }

  /**
   * Uploads a JSON file to a SharePoint document library
   */
  public async uploadJsonFile(
    libraryName: string,
    fileName: string,
    jsonContent: string
  ): Promise<string> {
    try {
      // Ensure the file has .json extension
      const sanitizedFileName = fileName.endsWith('.json') ? fileName : `${fileName}.json`;

      // Get the web URL
      const webUrl = this.context.pageContext.web.absoluteUrl;

      // Get the actual root folder path from the list
      const folderServerRelativeUrl = await this.getLibraryRootFolderServerRelativeUrl(libraryName);

      // Construct the endpoint URL for file upload
      const safeFolderUrl = this.escapeODataValue(folderServerRelativeUrl);
      const safeFileName = this.escapeODataValue(sanitizedFileName);

      const endpoint = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${safeFolderUrl}')/Files/add(url='${safeFileName}',overwrite=true)`;

      // Upload the file
      // Note: Let SPHttpClient handle headers automatically for file uploads
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body: jsonContent,
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to upload file: ${response.statusText}. ${errorText}`);
      }

      const result = await response.json();

      // Prefer ServerRelativeUrl for downstream metadata updates.
      if (result && result.ServerRelativeUrl) {
        return result.ServerRelativeUrl;
      }

      // Fallback
      return `${folderServerRelativeUrl}/${sanitizedFileName}`;
    } catch (error: any) {
      console.error('Error uploading JSON file:', error);
      throw error;
    }
  }

  /**
   * Lists files in the root of a document library
   */
  public async listLibraryFiles(libraryName: string): Promise<ISharePointFileInfo[]> {
    const page = await this.listLibraryFilesPage(libraryName, undefined, 5000);
    return page.items;
  }

  /**
   * Lists files in the root of a document library with paging.
   * Optionally expands ListItemAllFields to include column values.
   */
  public async listLibraryFilesPage<T extends ISharePointFileInfo = ISharePointFileInfo>(
    libraryName: string,
    fieldInternalNames?: string[],
    top: number = 50,
    nextLink?: string
  ): Promise<ISharePointFileInfoPage<T>> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;

      let endpoint: string;
      if (nextLink) {
        endpoint = nextLink;
      } else {
        const folderServerRelativeUrl = await this.getLibraryRootFolderServerRelativeUrl(libraryName);
        const safeFolderUrl = this.escapeODataValue(folderServerRelativeUrl);

        const baseSelect = ['Name', 'ServerRelativeUrl', 'TimeLastModified', 'Length'];

        const internalNames = (fieldInternalNames || []).filter(Boolean);
        const selectFields = internalNames.map((n) => `ListItemAllFields/${n}`);

        const select = [...baseSelect, ...selectFields].join(',');
        const expand = internalNames.length > 0 ? '&$expand=ListItemAllFields' : '';

        endpoint = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${safeFolderUrl}')/Files?$select=${select}${expand}&$top=${top}`;
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to list files: ${response.statusText}`);
      }

      const result = await response.json();
      const files = result?.value ?? [];

      const items = files.map((f: any) => ({
        name: f.Name,
        serverRelativeUrl: f.ServerRelativeUrl,
        timeLastModified: f.TimeLastModified,
        length: f.Length ? Number(f.Length) : undefined,
        ...(f.ListItemAllFields ? { ListItemAllFields: f.ListItemAllFields } : {}),
      }));

      const next = result['@odata.nextLink'] || result['odata.nextLink'] || result.__next;

      return { items: items as T[], nextLink: next };
    } catch (error: any) {
      console.error('Error listing library files (paged):', error);
      throw error;
    }
  }

  /**
   * Reads a file's contents as text by its ServerRelativeUrl
   */
  public async getFileTextByServerRelativeUrl(serverRelativeUrl: string): Promise<string> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const safeUrl = this.escapeODataValue(serverRelativeUrl);
      const endpoint = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${safeUrl}')/$value`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to download file: ${response.statusText}`);
      }

      return await response.text();
    } catch (error: any) {
      console.error('Error getting file text:', error);
      throw error;
    }
  }

  private escapeXmlAttribute(value: string): string {
    return (value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  private async loadLibraryFieldsCache(libraryName: string): Promise<void> {
    if (this._libraryFieldsCache[libraryName]) {
      return;
    }

    const webUrl = this.context.pageContext.web.absoluteUrl;
    const safeLibraryName = this.escapeODataValue(libraryName);

    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${safeLibraryName}')/fields?$select=InternalName&$top=5000`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to read list fields: ${response.statusText}`);
    }

    const result = await response.json();
    const fields = result?.value ?? [];

    const cache: { [internalName: string]: true } = {};
    for (const f of fields) {
      if (f && typeof f.InternalName === 'string') {
        cache[f.InternalName] = true;
      }
    }

    this._libraryFieldsCache[libraryName] = cache;
  }

  /**
   * Ensure that text columns exist in the library for the given internal names.
   */
  public async ensureTextFieldsExist(
    libraryName: string,
    fields: IEnsureTextFieldRequest[]
  ): Promise<void> {
    try {
      await this.loadLibraryFieldsCache(libraryName);

      const cache = this._libraryFieldsCache[libraryName];
      const missing = fields.filter((f) => f.internalName && !cache[f.internalName]);
      if (missing.length === 0) {
        return;
      }

      const webUrl = this.context.pageContext.web.absoluteUrl;
      const safeLibraryName = this.escapeODataValue(libraryName);
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${safeLibraryName}')/fields/createfieldasxml`;

      for (const f of missing) {
        const maxLength = f.maxLength || 255;

        const rawInternalName = f.internalName;
        const rawDisplayName = f.displayName || rawInternalName;

        const displayName = this.escapeXmlAttribute(rawDisplayName);
        const internalName = this.escapeXmlAttribute(rawInternalName);

        const schemaXml =
          `<Field Type="Text" DisplayName="${displayName}" Name="${internalName}" StaticName="${internalName}" Group="SmartForm" MaxLength="${maxLength}" />`;

        const body = JSON.stringify({
          parameters: {
            SchemaXml: schemaXml,
            Options: 0,
          },
        });

        const resp: SPHttpClientResponse = await this.context.spHttpClient.post(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
            },
            body,
          }
        );

        if (resp.ok) {
          cache[rawInternalName] = true;
        } else {
          // If it fails because it already exists or for permissions, continue.
          // We'll rely on the metadata update to fail/succeed and surface errors.
          const errorText = await resp.text();
          console.warn(`Failed to create field '${internalName}':`, errorText);
        }
      }
    } catch (error: any) {
      console.error('Error ensuring text fields:', error);
      throw error;
    }
  }

  /**
   * Updates a file's list item metadata.
   */
  public async updateFileListItemFields(
    fileServerRelativeUrl: string,
    fields: Record<string, any>
  ): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const safeUrl = this.escapeODataValue(fileServerRelativeUrl);

      const endpoint = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${safeUrl}')/ListItemAllFields`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify(fields),
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to update metadata: ${response.statusText}. ${errorText}`);
      }
    } catch (error: any) {
      console.error('Error updating file metadata:', error);
      throw error;
    }
  }

  /**
   * Deletes a file by its ServerRelativeUrl
   */
  public async deleteFileByServerRelativeUrl(serverRelativeUrl: string): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const safeUrl = this.escapeODataValue(serverRelativeUrl);
      const endpoint = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${safeUrl}')`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
          },
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to delete file: ${response.statusText}. ${errorText}`);
      }
    } catch (error: any) {
      console.error('Error deleting file:', error);
      throw error;
    }
  }

  /**
   * Gets the current user's information
   */
  public async getCurrentUser(): Promise<ISharePointUser> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const endpoint = `${webUrl}/_api/web/currentuser`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to get current user: ${response.statusText}`);
      }

      const user = await response.json();

      return {
        displayName: user.Title || 'Unknown User',
        email: user.Email || '',
        loginName: user.LoginName || '',
      };
    } catch (error) {
      console.error('Error getting current user:', error);
      // Return fallback user info from page context
      return {
        displayName: this.context.pageContext.user.displayName || 'Unknown User',
        email: this.context.pageContext.user.email || '',
        loginName: this.context.pageContext.user.loginName || '',
      };
    }
  }

  /**
   * Checks if a document library exists
   */
  public async ensureLibraryExists(libraryName: string): Promise<boolean> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const safeLibraryName = this.escapeODataValue(libraryName);
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${safeLibraryName}')`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      return response.ok;
    } catch (error) {
      console.error('Error checking library existence:', error);
      return false;
    }
  }

  /**
   * Checks if the current user has write permission to a library
   */
  public async checkWritePermission(libraryName: string): Promise<boolean> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const safeLibraryName = this.escapeODataValue(libraryName);
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${safeLibraryName}')/effectiveBasePermissions`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const permissions = await response.json();

      // Check for AddListItems permission (value: 2)
      // High and Low represent the permission mask
      const addItemsPermission = 2;
      const hasPermission = (permissions.Low & addItemsPermission) === addItemsPermission;

      return hasPermission;
    } catch (error) {
      console.error('Error checking write permission:', error);
      return false;
    }
  }

  /**
   * Searches for SharePoint users (for People Picker)
   */
  public async searchUsers(searchText: string, maxResults: number = 10): Promise<any[]> {
    try {
      if (!searchText || searchText.trim().length === 0) {
        return [];
      }

      const webUrl = this.context.pageContext.web.absoluteUrl;
      const filter = `substringof('${searchText}', Title) or substringof('${searchText}', Email)`;
      const endpoint = `${webUrl}/_api/web/siteusers?$filter=${filter}&$top=${maxResults}&$select=Id,Title,Email,LoginName`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to search users: ${response.statusText}`);
      }

      const result = await response.json();
      return result.value || [];
    } catch (error) {
      console.error('Error searching users:', error);
      return [];
    }
  }
}
