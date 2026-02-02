/**
 * Service for Microsoft Graph API operations
 * Handles user profile and manager lookup
 */

import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/photos';
import '@pnp/graph/groups';
import '@pnp/graph/members';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserInfo, IManagerInfo } from '../models/IAbsenceRegistration';

/**
 * Interface for group member information
 */
export interface IGroupMember {
  id: string;
  displayName: string;
  email: string;
  jobTitle?: string;
}

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
    } catch {
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
    } catch {
      // User might not have a manager
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
    } catch {
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
    } catch {
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

  /**
   * Get members of a security group
   * @param groupId The Azure AD group ID
   */
  public async getGroupMembers(groupId: string): Promise<IGroupMember[]> {
    try {
      const members = await this._graph.groups.getById(groupId).members();

      return members
        .filter((member) => (member as { '@odata.type'?: string })['@odata.type'] === '#microsoft.graph.user')
        .map((member) => ({
          id: (member as { id?: string }).id || '',
          displayName: (member as { displayName?: string }).displayName || '',
          email: (member as { mail?: string }).mail || (member as { userPrincipalName?: string }).userPrincipalName || '',
          jobTitle: (member as { jobTitle?: string }).jobTitle,
        }));
    } catch {
      return [];
    }
  }

  /**
   * Add a user to a security group
   * @param groupId The Azure AD group ID
   * @param userId The Azure AD user ID
   */
  public async addUserToGroup(groupId: string, userId: string): Promise<void> {
    await this._graph.groups.getById(groupId).members.add(
      `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
    );
  }

  /**
   * Remove a user from a security group
   * @param groupId The Azure AD group ID
   * @param userId The Azure AD user ID
   */
  public async removeUserFromGroup(groupId: string, userId: string): Promise<void> {
    await this._graph.groups.getById(groupId).members.getById(userId).remove();
  }

  /**
   * Search users and return with their Azure AD ID (needed for group operations)
   * @param searchText Search query
   */
  public async searchUsersWithId(searchText: string): Promise<IGroupMember[]> {
    if (!searchText || searchText.length < 2) {
      return [];
    }

    try {
      const users = await this._graph.users
        .filter(`startswith(displayName,'${searchText}') or startswith(mail,'${searchText}') or startswith(userPrincipalName,'${searchText}')`)
        .select('id', 'displayName', 'mail', 'userPrincipalName', 'jobTitle')
        .top(10)();

      return users.map((user) => ({
        id: user.id || '',
        displayName: user.displayName || '',
        email: user.mail || user.userPrincipalName || '',
        jobTitle: user.jobTitle || undefined,
      }));
    } catch {
      return [];
    }
  }
}

export default GraphService;
