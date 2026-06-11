export interface IGroupSummary {
    id?: number | string;
    name: string;
    userCount?: number;
    possibleExternalUserCount?: number;
    isEmpty?: boolean;
    
    /** Indicates if this group contains 'Everyone' or 'Everyone except external users' claims (Copilot risk) */
    isOvershared?: boolean;

    /** Additional context or warnings about this group */
    notes?: string[];
}
