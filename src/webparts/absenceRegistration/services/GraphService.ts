/**
 * Service for Microsoft Graph API operations
 * Handles user profile and manager lookup
 */

import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/photos';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserInfo, IManagerInfo } from '../models/IAbsenceRegistration';

export class GraphService {
  private static _instance: GraphService;
  private _graph: ReturnType<typeof graphfi>;
  private _context: WebPartContext;

  private constructor() {
    // Private constructor for singleton
  }

  /**
   * Initialize the Graph service with SPFx context
   */
  public static getInstance(context?: WebPartContext): GraphService {
    if (!GraphService._instance) {
      GraphService._instance = new GraphService();
    }

    if (context) {
      GraphService._instance._context = context;
      GraphService._instance._graph = graphfi().using(graphSPFx(context));
    }

    return GraphService._instance;
  }

  /**
   * Get the current user's profile information
   */
  public async getCurrentUser(): Promise<IUserInfo> {
    try {
      const user = await this._graph.me();

      return {
        id: 0, // Will be resolved via SharePoint
        email: user.mail || user.userPrincipalName || '',
        displayName: user.displayName || '',
        department: user.department || '',
        jobTitle: user.jobTitle || '',
      };
    } catch (error) {
      console.error('Error fetching current user:', error);
      throw new Error('Kunne ikke hente brugeroplysninger');
    }
  }

  /**
   * Get the current user's manager
   */
  public async getCurrentUserManager(): Promise<IManagerInfo | undefined> {
    try {
      const manager = await this._graph.me.manager();

      if (manager) {
        return {
          id: (manager as { id?: string }).id || '',
          email:
            (manager as { mail?: string }).mail ||
            (manager as { userPrincipalName?: string }).userPrincipalName ||
            '',
          displayName: (manager as { displayName?: string }).displayName || '',
        };
      }

      return undefined;
    } catch (error) {
      // User might not have a manager
      console.warn('No manager found or error fetching manager:', error);
      return undefined;
    }
  }

  /**
   * Get user information by email
   */
  public async getUserByEmail(email: string): Promise<IUserInfo | undefined> {
    try {
      const users = await this._graph.users
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id', 'displayName', 'mail', 'userPrincipalName', 'department')
        .top(1)();

      if (users && users.length > 0) {
        const user = users[0];
        return {
          id: 0,
          email: user.mail || user.userPrincipalName || '',
          displayName: user.displayName || '',
          department: user.department || '',
        };
      }

      return undefined;
    } catch (error) {
      console.error('Error fetching user by email:', error);
      return undefined;
    }
  }

  /**
   * Search users by name or email (for people picker)
   */
  public async searchUsers(searchText: string): Promise<IUserInfo[]> {
    if (!searchText || searchText.length < 2) {
      return [];
    }

    try {
      const users = await this._graph.users
        .filter(`startswith(displayName,'${searchText}') or startswith(mail,'${searchText}') or startswith(userPrincipalName,'${searchText}')`)
        .select('id', 'displayName', 'mail', 'userPrincipalName', 'jobTitle')
        .top(10)();

      return users.map((user) => ({
        id: 0,
        email: user.mail || user.userPrincipalName || '',
        displayName: user.displayName || '',
        jobTitle: user.jobTitle || '',
      }));
    } catch (error) {
      console.error('Error searching users:', error);
      return [];
    }
  }

  /**
   * Get user's photo as base64 (optional feature)
   */
  public async getUserPhoto(): Promise<string | undefined> {
    try {
      const photo = await this._graph.me.photo.getBlob();
      return URL.createObjectURL(photo);
    } catch {
      // User might not have a photo
      return undefined;
    }
  }

  /**
   * Get full user info including manager
   */
  public async getFullUserInfo(): Promise<IUserInfo> {
    const user = await this.getCurrentUser();
    const manager = await this.getCurrentUserManager();

    return {
      ...user,
      manager,
    };
  }
}

export default GraphService;
