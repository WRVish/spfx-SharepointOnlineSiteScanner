import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/security/list';
import '@pnp/sp/items';

import { IPermissionIndicator } from '../models/IPermissionIndicator';
import { MAX_LIBRARIES_TO_SCAN } from '../constants/RiskConstants';
import { LoggerService } from './LoggerService';

const SOURCE = 'PermissionService';

export class PermissionService {
    private _sp: SPFI;

    constructor(sp: SPFI) {
        this._sp = sp;
    }

    /**
     * Gets permission indicators for the current site and its document libraries.
     * Combines site-level and library-level checks.
     */
    public async getPermissionIndicators(): Promise<IPermissionIndicator[]> {
        LoggerService.info(SOURCE, 'Gathering permission indicators...');

        const indicators: IPermissionIndicator[] = [];

        // Check if the Site itself is Overshared (Copilot risk)
        try {
            const siteIndicator = await this.getSitePermissionIndicator();
            if (siteIndicator) {
                indicators.push(siteIndicator);
            }
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to check site oversharing', error);
        }

        // Get library-level permission indicators
        try {
            const libraryIndicators = await this.getLibraryPermissionIndicators();
            indicators.push(...libraryIndicators);
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to check library permissions', error);
            indicators.push({
                scope: 'Library',
                title: 'Document Libraries',
                status: 'Error',
                notes: ['Unable to check library permission inheritance.'],
            });
        }

        return indicators;
    }

    /**
     * Checks the site strictly for Copilot Oversharing ("Everyone" claims).
     */
    private async getSitePermissionIndicator(): Promise<IPermissionIndicator | null> {
        try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const webInfo: any = await this._sp.web.select('Title', 'Url')();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const roleAssignments: any[] = await this._sp.web.roleAssignments.expand("Member")();
            
            const isOvershared = roleAssignments.some(ra => {
                const loginName = ra.Member?.LoginName?.toLowerCase() || '';
                return loginName.includes('spo-grid-all-users') || loginName === 'everyone';
            });

            if (isOvershared) {
                return {
                    scope: 'Site',
                    title: webInfo.Title || 'Current Site',
                    url: webInfo.Url,
                    hasUniquePermissions: false, // We ignore site unique perms for noise reduction
                    status: 'Review',
                    notes: ['CRITICAL: The entire site is overshared with Everyone or All Company. Copilot may expose this data.'],
                    isOvershared: true,
                    anonymousLinkCount: 0
                };
            }
        } catch (err) {
            LoggerService.error(SOURCE, 'Failed to check site oversharing', err);
        }
        return null;
    }

    /**
     * Checks document libraries (BaseTemplate 101, non-hidden) for unique permissions.
     * Limited to MAX_LIBRARIES_TO_SCAN to prevent performance issues.
     */
    public async getLibraryPermissionIndicators(): Promise<IPermissionIndicator[]> {
        LoggerService.info(SOURCE, 'Checking document library permissions...');

        try {
            // Get non-hidden document libraries
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const lists: any[] = await this._sp.web.lists
                .filter('BaseTemplate eq 101 and Hidden eq false')
                .select('Id', 'Title', 'RootFolder/ServerRelativeUrl', 'HasUniqueRoleAssignments', 'ItemCount', 'LastItemModifiedDate')
                .expand('RootFolder')
                .top(MAX_LIBRARIES_TO_SCAN)();

            LoggerService.info(SOURCE, `Found ${lists.length} document libraries to check`);

            const indicators: IPermissionIndicator[] = [];

            for (const list of lists) {
                let isOvershared = false;
                let anonymousLinkCount = 0;
                const notes: string[] = [];

                if (list.HasUniqueRoleAssignments) {
                    notes.push('This library has unique permissions — access may differ from site defaults.');
                    
                    try {
                        // Check for oversharing ("Everyone", "Everyone except external users")
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const roleAssignments: any[] = await this._sp.web.lists.getById(list.Id).roleAssignments.expand("Member")();
                        isOvershared = roleAssignments.some(ra => {
                            const loginName = ra.Member?.LoginName?.toLowerCase() || '';
                            return loginName.includes('spo-grid-all-users') || loginName === 'everyone';
                        });

                        if (isOvershared) {
                            notes.push('CRITICAL: This library is overshared with Everyone or All Company.');
                        }

                        // Check for anonymous links on items (limited to top 10 unique items to prevent timeouts)
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const itemsWithUniquePerms: any[] = await this._sp.web.lists.getById(list.Id).items.filter('HasUniqueRoleAssignments eq true').select('Id', 'FileLeafRef', 'Title').expand('File').top(10)();
                        
                        for (const item of itemsWithUniquePerms) {
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            const itemRoles: any[] = await this._sp.web.lists.getById(list.Id).items.getById(item.Id).roleAssignments.expand("Member")();
                            const hasAnonLink = itemRoles.some(ra => ra.Member?.LoginName?.toLowerCase().includes('urn:spo:anon'));
                            if (hasAnonLink) {
                                indicators.push({
                                    scope: 'Item',
                                    title: item.FileLeafRef || item.Title || `Item ID: ${item.Id} (in ${list.Title})`,
                                    url: `${list.RootFolder?.ServerRelativeUrl}/DispForm.aspx?ID=${item.Id}`,
                                    hasUniquePermissions: true,
                                    status: 'Review',
                                    notes: ['Active anonymous link detected'],
                                    anonymousLinkCount: 1
                                });
                            }
                        }

                        // We don't push library-level notes for anon links anymore, the Item-level indicators handle it.

                    } catch (err) {
                        LoggerService.error(SOURCE, `Failed to check detailed permissions for list ${list.Title}`, err);
                    }
                }

                indicators.push({
                    scope: 'Library',
                    title: list.Title,
                    url: list.RootFolder?.ServerRelativeUrl,
                    hasUniquePermissions: list.HasUniqueRoleAssignments,
                    itemCount: list.ItemCount,
                    lastModified: list.LastItemModifiedDate,
                    status: list.HasUniqueRoleAssignments ? 'Review' : 'Ok',
                    notes: notes.length > 0 ? notes : undefined,
                    isOvershared,
                    anonymousLinkCount: 0 // Zeroed out so we don't double count the item-level indicators
                });
            }

            return indicators;
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to enumerate libraries', error);
            throw error;
        }
    }
}
