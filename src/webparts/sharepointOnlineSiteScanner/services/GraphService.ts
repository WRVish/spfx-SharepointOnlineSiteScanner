import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LoggerService } from './LoggerService';

const SOURCE = 'GraphService';

export interface IGuestGraphInfo {
    id: string;
    displayName: string;
    userPrincipalName: string;
    userType: string;
    lastSignInDateTime?: string;
}

export class GraphService {
    private _context: WebPartContext;
    private _graphClient: MSGraphClientV3 | null = null;

    constructor(context: WebPartContext) {
        this._context = context;
    }

    private async getClient(): Promise<MSGraphClientV3> {
        if (!this._graphClient) {
            this._graphClient = await this._context.msGraphClientFactory.getClient('3');
        }
        return this._graphClient;
    }

    /**
     * Fetches all Guest users from the tenant's Entra ID along with their last sign-in activity.
     */
    public async getTenantGuests(): Promise<IGuestGraphInfo[]> {
        try {
            LoggerService.info(SOURCE, 'Fetching tenant guests from Graph API...');
            const client = await this.getClient();
            // Note: signInActivity requires AuditLog.Read.All
            const response = await client.api('/users')
                .select('id,displayName,userPrincipalName,userType,signInActivity')
                .filter("userType eq 'Guest'")
                .get();
                
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const guests = response.value.map((u: any) => ({
                id: u.id,
                displayName: u.displayName,
                userPrincipalName: u.userPrincipalName,
                userType: u.userType,
                lastSignInDateTime: u.signInActivity?.lastSignInDateTime
            }));
            
            LoggerService.info(SOURCE, `Found ${guests.length} tenant guests`);
            return guests;
        } catch (error) {
            LoggerService.error(SOURCE, 'Error fetching tenant guests via Graph API. Check API permissions.', error);
            return [];
        }
    }
}
