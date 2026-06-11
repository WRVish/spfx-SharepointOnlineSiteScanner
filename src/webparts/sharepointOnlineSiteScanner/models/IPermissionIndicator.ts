export interface IPermissionIndicator {
    scope: 'Site' | 'Library' | 'Folder' | 'Item' | 'Unknown';
    title: string;
    url?: string;
    hasUniquePermissions?: boolean;
    status: 'Ok' | 'Review' | 'Unknown' | 'Error';
    itemCount?: number;
    lastModified?: string;
    notes?: string[];
    isOvershared?: boolean;
    anonymousLinkCount?: number;
}
