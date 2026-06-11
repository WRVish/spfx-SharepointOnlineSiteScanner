import * as React from 'react';
import { SPFI } from '@pnp/sp';
import { Icon } from '@fluentui/react/lib/Icon';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Link } from '@fluentui/react/lib/Link';

import styles from './SharepointOnlineSiteScanner.module.scss';
import { LoadingState } from './LoadingState';

import { IScanResult } from '../models/IScanResult';
import { RiskLabel } from '../models/IRiskScore';
import { IGroupSummary } from '../models/IGroupSummary';
import { IPermissionIndicator } from '../models/IPermissionIndicator';
import { IUserRiskInfo } from '../models/IUserRiskInfo';

import { GraphService } from '../services/GraphService';
import { SharePointSiteService } from '../services/SharePointSiteService';
import { SharePointUserService } from '../services/SharePointUserService';
import { SharePointGroupService } from '../services/SharePointGroupService';
import { PermissionService } from '../services/PermissionService';
import { RiskScoringService } from '../services/RiskScoringService';
import { LoggerService } from '../services/LoggerService';

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ModernDashboard } from './ModernDashboard';
import { CompactDashboard } from './CompactDashboard';

export type UITheme = 'classic' | 'modern' | 'compact' | 'unified';

// ═══════════════════════════════════
// Props
// ═══════════════════════════════════

export interface ISharepointOnlineSiteScannerProps {
    sp: SPFI;
    context: WebPartContext;
    internalDomains: string;
    siteUrl: string;
    uiTheme: UITheme;
    customCaps: any;
}

// ═══════════════════════════════════
// State
// ═══════════════════════════════════

interface IComponentState {
    phase: 'idle' | 'scanning' | 'done' | 'error';
    scanResult: IScanResult | null;
    errorMessage: string;
}

// ═══════════════════════════════════
// Component
// ═══════════════════════════════════

export class SharepointOnlineSiteScanner extends React.Component<
    ISharepointOnlineSiteScannerProps,
    IComponentState
> {
    constructor(props: ISharepointOnlineSiteScannerProps) {
        super(props);
        this.state = { phase: 'idle', scanResult: null, errorMessage: '' };
    }

    // ── Scanning Logic ──

    private async runScan(): Promise<void> {
        this.setState({ phase: 'scanning', errorMessage: '' });

        try {
            const { sp, internalDomains, siteUrl } = this.props;
            const domainList = internalDomains
                ? internalDomains.split(',').map((d) => d.trim().toLowerCase()).filter(Boolean)
                : [];

            const siteService = new SharePointSiteService(sp);
            const userService = new SharePointUserService(sp);
            const groupService = new SharePointGroupService(sp);
            const permissionService = new PermissionService(sp);
            const riskService = new RiskScoringService();
            const graphService = new GraphService(this.props.context);

            LoggerService.info('Scanner', `Starting scan for ${siteUrl}`);

            const siteSummary = await siteService.getSiteSummary();
            const users = await userService.getSiteUsers(domainList);
            const groups = await groupService.getSiteGroups(domainList);
            const permissionIndicators = await permissionService.getPermissionIndicators();
            const tenantGuests = await graphService.getTenantGuests();
            
            // Cross-reference users with true Entra ID guests
            if (tenantGuests.length > 0) {
                users.forEach(u => {
                    const matchedGuest = tenantGuests.find(g => 
                        g.userPrincipalName.toLowerCase() === u.loginName?.toLowerCase() || 
                        g.userPrincipalName.toLowerCase().includes(u.email?.toLowerCase() || 'NO_MATCH')
                    );
                    if (matchedGuest) {
                        u.isPossibleExternal = true;
                        if (!u.detectionReasons.includes('Verified via Entra ID')) {
                            u.detectionReasons.push('Verified via Entra ID (Graph API)');
                        }
                    }
                });
            }

            const partialResult = {
                siteSummary,
                users,
                groups,
                permissionIndicators,
                warnings: [] as string[],
                errors: [] as any[],
            };

            const riskScore = riskService.calculateRiskScore(partialResult, this.props.customCaps);

            const scanResult: IScanResult = {
                ...partialResult,
                riskScore,
            };

            LoggerService.info('Scanner', `Scan complete. Risk: ${riskScore.label} (${riskScore.score})`);
            this.setState({ phase: 'done', scanResult });
        } catch (err: any) {
            LoggerService.error('Scanner', 'Scan failed', err);
            this.setState({ phase: 'error', errorMessage: err.message || 'Unknown error' });
        }
    }

    // ── Render ──

    public render(): React.ReactElement {
        const { phase, scanResult, errorMessage } = this.state;

        if (phase === 'scanning') {
            return <LoadingState />;
        }

        if (phase === 'error') {
            return (
                <div className={styles.dashboardRoot}>
                    <div className={styles.topBar}>
                        <h1 className={styles.topBarTitle}>Sharepoint-Online-Site-Scanner</h1>
                        <button className={styles.scanButton} onClick={() => this.runScan()}>
                            <Icon iconName="Refresh" /> Retry
                        </button>
                    </div>
                    <div className={styles.warningBanner}>
                        <Icon iconName="ErrorBadge" className={styles.warningBannerIcon} />
                        <span>Scan failed: {errorMessage}</span>
                    </div>
                </div>
            );
        }

        if (phase === 'idle' || !scanResult) {
            return this.renderPreScan();
        }

        if (this.props.uiTheme === 'modern') {
            return (
                <div>
                    <ModernDashboard 
                        scanResult={scanResult} 
                        onRescan={() => this.runScan()} 
                        onExportCsv={() => this.exportToCsv(scanResult)}
                    />
                    {/* Keep detailed view available below for deep dive */}
                    <div style={{ marginTop: '20px', padding: '0 24px' }}>
                        <h2 style={{fontSize: '18px'}}>Deep Dive (Classic Detailed View)</h2>
                        {this.renderDetailedView(scanResult)}
                    </div>
                </div>
            );
        }

        if (this.props.uiTheme === 'compact') {
            return (
                <div>
                    <CompactDashboard 
                        scanResult={scanResult}
                        onExportCsv={() => this.exportToCsv(scanResult)}
                    />
                </div>
            );
        }

        if (this.props.uiTheme === 'unified') {
            return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '30px' }}>
                    <div id="unified-summary" style={{ background: '#ffffff', padding: '10px' }}>
                        {this.renderDashboard(scanResult, true)}
                    </div>

                    <div id="unified-external" style={{ padding: '20px 0', borderTop: '2px solid #e1dfdd' }}>
                        <h2 style={{ fontSize: '20px', marginBottom: '15px' }}>External Users</h2>
                        {this.renderExternalUsersTable(scanResult.users.filter(u => u.isPossibleExternal), scanResult.siteSummary.url)}
                    </div>

                    <div id="unified-groups" style={{ padding: '20px 0', borderTop: '2px solid #e1dfdd' }}>
                        <h2 style={{ fontSize: '20px', marginBottom: '15px' }}>Groups with External Users</h2>
                        {this.renderGroupsWithExternalTable(scanResult.groups.filter(g => (g.possibleExternalUserCount ?? 0) > 0), scanResult.siteSummary.url)}
                    </div>

                    <div id="unified-perms" style={{ padding: '20px 0', borderTop: '2px solid #e1dfdd' }}>
                        <h2 style={{ fontSize: '20px', marginBottom: '15px' }}>Resources with Unique Permissions</h2>
                        {this.renderUniquePermissionsTable(scanResult.permissionIndicators.filter(p => p.hasUniquePermissions), scanResult.siteSummary.url)}
                    </div>

                    <div id="unified-anon-links" style={{ padding: '20px 0', borderTop: '2px solid #e1dfdd' }}>
                        <h2 style={{ fontSize: '20px', marginBottom: '15px' }}>Resources with Anonymous Links</h2>
                        {this.renderAnonymousLinksTable(scanResult.permissionIndicators.filter(p => (p.anonymousLinkCount ?? 0) > 0), scanResult.siteSummary.url)}
                    </div>

                    {scanResult.riskScore.recommendedActions.length > 0 && (
                        <div id="unified-actions" style={{ padding: '20px 0', borderTop: '2px solid #e1dfdd' }}>
                            {this.renderActionsCard(scanResult.riskScore.recommendedActions)}
                        </div>
                    )}
                </div>
            );
        }

        return (
            <Pivot aria-label="Risk Scanner Views">
                <PivotItem headerText="Risk Dashboard">
                    {this.renderDashboard(scanResult)}
                </PivotItem>
                <PivotItem headerText="Detailed Risk View">
                    {this.renderDetailedView(scanResult)}
                </PivotItem>
            </Pivot>
        );
    }

    // ── Pre-scan state ──

    private renderPreScan(): React.ReactElement {
        return (
            <div className={styles.dashboardRoot}>
                <div className={styles.topBar}>
                    <h1 className={styles.topBarTitle}>Sharepoint-Online-Site-Scanner</h1>
                    <button
                        className={styles.scanButton}
                        onClick={() => this.runScan()}
                    >
                        <Icon iconName="Shield" /> Run Scan
                    </button>
                </div>
                <div className={styles.preScanAlert}>
                    <Icon iconName="Info" className={styles.preScanIcon} />
                    <span>
                        Click <strong>Run Scan</strong> to analyze this site&#39;s external sharing configuration,
                        user permissions, and group memberships for potential risk signals.
                    </span>
                </div>
            </div>
        );
    }

    // ── Dashboard (post-scan) ──

    private renderDashboard(result: IScanResult, isUnified: boolean = false): React.ReactElement {
        const { riskScore, users, groups, permissionIndicators, siteSummary } = result;
        const externalUsers = users.filter((u) => u.isPossibleExternal);
        const uniquePermCount = permissionIndicators.filter((p) => p.hasUniquePermissions).length;

        return (
            <div className={styles.dashboardRoot}>
                {/* Top bar with title + rescan */}
                <div className={styles.topBar}>
                    <h1 className={styles.topBarTitle}>Sharepoint-Online-Site-Scanner</h1>
                    <button
                        className={styles.scanButton}
                        onClick={() => this.runScan()}
                    >
                        <Icon iconName="Refresh" /> Re-scan
                    </button>
                </div>

                {/* ── Metric Cards ── */}
                <div style={{ paddingTop: '20px' }}>
                    <div className={styles.metricGrid}>
                    {this.renderMetricCard(
                        'Shield',
                        'Risk Score',
                        `${riskScore.score}/100`,
                        riskScore.label,
                        this.getRiskStatusClass(riskScore.score),
                        true
                    )}
                    {this.renderMetricCard(
                        'People',
                        'External Users',
                        String(externalUsers.length),
                        externalUsers.length === 0 ? 'None detected' : `${externalUsers.length} found`,
                        externalUsers.length === 0 ? 'Success' : 'Warning'
                    )}
                    {this.renderMetricCard(
                        'Group',
                        'Groups Reviewed',
                        String(groups.length),
                        `${groups.filter((g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0).length} with external`,
                        groups.filter((g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0).length > 0 ? 'Warning' : 'Success'
                    )}
                    {this.renderMetricCard(
                        'Lock',
                        'Unique Permissions',
                        String(uniquePermCount),
                        uniquePermCount > 0 ? `${uniquePermCount} detected` : 'Inheriting normally',
                        uniquePermCount > 0 ? 'Warning' : 'Success'
                    )}
                    {this.renderMetricCard(
                        'DocumentSet',
                        'Libraries Reviewed',
                        String(permissionIndicators.filter((p) => p.scope === 'Library').length),
                        'Scanned',
                        'Info'
                    )}
                </div>

                {/* ── Last Scanned ── */}
                <div className={styles.lastScannedRow}>
                    <Icon iconName="Clock" className={styles.lastScannedIcon} />
                    <span>Last scanned: {this.formatDate(siteSummary.scannedAt)}</span>
                </div>

                {/* ── Warnings ── */}
                {result.warnings.length > 0 && (
                    <div className={styles.warningBanner}>
                        <Icon iconName="Warning" className={styles.warningBannerIcon} />
                        <span>{result.warnings.join(' | ')}</span>
                    </div>
                )}

                {/* ── Main Assessment Grid: Risk + Key Findings ── */}
                <div className={isUnified ? '' : styles.mainAssessmentGrid} style={isUnified ? { display: 'flex', flexDirection: 'column', gap: '24px' } : {}}>
                    {this.renderRiskAssessmentCard(result, isUnified)}
                    {this.renderKeyFindingsCard(result)}
                </div>

                {/* ── Secondary Grid: Groups + Permissions + External Indicators ── */}
                {!isUnified && (
                    <div className={styles.secondaryGrid}>
                        {this.renderGroupsCard(groups)}
                        {this.renderPermissionsCard(permissionIndicators)}
                        {this.renderExternalIndicatorsCard(externalUsers)}
                    </div>
                )}

                {/* ── Recommended Actions ── */}
                {!isUnified && riskScore.recommendedActions.length > 0 && this.renderActionsCard(riskScore.recommendedActions)}
            </div>
                
                {/* ── Footer ── */}
                <div className={styles.dashboardFooter}>
                    Sharepoint-Online-Site-Scanner v1.0 &mdash;{' '}
                    <a
                        href="https://github.com/BillySharePoint/spfx-external-sharing-risk-scanner"
                        target="_blank"
                        rel="noopener noreferrer"
                        className={styles.footerLink}
                    >
                        GitHub
                    </a>
                </div>
            </div>
        );
    }

    
        // ── Export ──
    private exportToCsv(result: IScanResult): void {
        const rows = [
            ['Report', 'SharePoint Site Scanner'],
            ['Site URL', result.siteSummary.url],
            ['Risk Score', `${result.riskScore.score} (${result.riskScore.label})`],
            [],
            ['--- EXTERNAL USERS ---'],
            ['Display Name', 'Email', 'Login Name', 'Reasons']
        ];
        
        result.users.filter(u => u.isPossibleExternal).forEach(u => {
            rows.push([u.displayName, u.email || '', u.loginName || '', u.detectionReasons.join(', ')]);
        });
        
        rows.push([], ['--- GROUPS WITH EXTERNAL USERS ---'], ['Group Name', 'Total Users', 'External Users']);
        result.groups.filter(g => (g.possibleExternalUserCount ?? 0) > 0).forEach(g => {
            rows.push([g.name, String(g.userCount), String(g.possibleExternalUserCount)]);
        });
        
        const csvContent = rows.map(e => e.join(",")).join("\n");
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", "Site_Risk_Report.csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }


    
    // ── Detailed View ──

    private renderDetailedView(result: IScanResult): React.ReactElement {
        const { users, groups, permissionIndicators, siteSummary } = result;
        const externalUsers = users.filter((u) => u.isPossibleExternal);
        const uniquePerms = permissionIndicators.filter((p) => p.hasUniquePermissions);
        const anonLinks = permissionIndicators.filter((p) => (p.anonymousLinkCount ?? 0) > 0);
        const groupsWithExt = groups.filter((g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0);
        const emptyGroups = groups.filter((g) => g.isEmpty);

        return (
            <div className={styles.detailedViewWrap}>
                <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '10px', gap: '10px' }}>
                    <button className={styles.scanButton} onClick={() => this.exportToCsv(result)}>
                        <Icon iconName="DownloadDocument" /> Export to CSV
                    </button>
                </div>
                {this.renderExternalUsersTable(externalUsers, siteSummary.url)}

                <div style={{ marginTop: '20px' }}>
                    {this.renderUniquePermissionsTable(uniquePerms, siteSummary.url)}
                </div>

                <div style={{ marginTop: '20px' }}>
                    {this.renderAnonymousLinksTable(anonLinks, siteSummary.url)}
                </div>

                {this.renderGroupsWithExternalTable(groupsWithExt, siteSummary.url)}
                {this.renderEmptyGroupsTable(emptyGroups, siteSummary.url)}
            </div>
        );
    }

    private renderExternalUsersTable(users: IUserRiskInfo[], siteUrl: string): React.ReactElement {
        const columns: IColumn[] = [
            { key: 'name', name: 'Display Name', fieldName: 'displayName', minWidth: 150, maxWidth: 250, isResizable: true },
            { key: 'email', name: 'Email', fieldName: 'email', minWidth: 200, maxWidth: 300, isResizable: true },
            { key: 'login', name: 'Login Name', fieldName: 'loginName', minWidth: 200, isResizable: true },
            { 
                key: 'action', 
                name: 'Action', 
                minWidth: 100, 
                onRender: () => (
                    <Link href={`${siteUrl}/_layouts/15/people.aspx`} target="_blank">Manage Users</Link>
                )
            }
        ];

        return (
            <div className={`${styles.card} ${styles.detailedCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="PeopleAlert" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>External Users Roster ({users.length})</h2>
                </div>
                <div className={styles.cardBody}>
                    {users.length > 0 ? (
                        <DetailsList
                            items={users}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                        />
                    ) : (
                        <p>No external users found.</p>
                    )}
                </div>
            </div>
        );
    }

    private renderGroupsWithExternalTable(groups: IGroupSummary[], siteUrl: string): React.ReactElement {
        const columns: IColumn[] = [
            { key: 'name', name: 'Group Name', fieldName: 'name', minWidth: 200, maxWidth: 300, isResizable: true },
            { key: 'total', name: 'Total Users', fieldName: 'userCount', minWidth: 100, maxWidth: 100 },
            { key: 'ext', name: 'External Users', fieldName: 'possibleExternalUserCount', minWidth: 100, maxWidth: 100 },
            { 
                key: 'action', 
                name: 'Action', 
                minWidth: 150, 
                onRender: (item: IGroupSummary) => (
                    <Link href={`${siteUrl}/_layouts/15/people.aspx?MembershipGroupId=${item.id}`} target="_blank">View Membership</Link>
                )
            }
        ];

        return (
            <div className={`${styles.card} ${styles.detailedCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Group" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Groups Exposing External Access ({groups.length})</h2>
                </div>
                <div className={styles.cardBody}>
                    {groups.length > 0 ? (
                        <DetailsList
                            items={groups}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                        />
                    ) : (
                        <p>No groups contain external users.</p>
                    )}
                </div>
            </div>
        );
    }

    private renderUniquePermissionsTable(permissions: IPermissionIndicator[], siteUrl: string): React.ReactElement {
        const columns: IColumn[] = [
            { 
                key: 'title', 
                name: 'Resource Name', 
                fieldName: 'title', 
                minWidth: 150, 
                maxWidth: 250, 
                isResizable: true,
                onRender: (item: IPermissionIndicator) => `${item.title} (${item.scope})`
            },
            { 
                key: 'overshared', 
                name: 'Overshared?', 
                minWidth: 90, 
                maxWidth: 100,
                onRender: (item: IPermissionIndicator) => (
                    <span style={{ color: item.isOvershared ? '#d83b01' : 'inherit', fontWeight: item.isOvershared ? 'bold' : 'normal' }}>
                        {item.isOvershared ? 'Yes (Critical)' : 'No'}
                    </span>
                )
            },
            { 
                key: 'anonLinks', 
                name: 'Anon Links', 
                minWidth: 80, 
                maxWidth: 100,
                onRender: (item: IPermissionIndicator) => (
                    <span style={{ color: (item.anonymousLinkCount || 0) > 0 ? '#d83b01' : 'inherit', fontWeight: (item.anonymousLinkCount || 0) > 0 ? 'bold' : 'normal' }}>
                        {item.anonymousLinkCount || 0}
                    </span>
                )
            },
            { 
                key: 'items', 
                name: 'Items', 
                fieldName: 'itemCount', 
                minWidth: 60, 
                maxWidth: 80,
                onRender: (item: IPermissionIndicator) => item.scope === 'Site' ? '—' : item.itemCount 
            },
            { 
                key: 'action', 
                name: 'Action', 
                minWidth: 120, 
                onRender: (item: IPermissionIndicator) => {
                    if (item.scope === 'Site') {
                        return <Link href={`${siteUrl}/_layouts/15/user.aspx`} target="_blank">Site Permissions</Link>;
                    }
                    return <Link href={`${siteUrl}${item.url ? item.url : ''}/Forms/AllItems.aspx`} target="_blank">View Library</Link>;
                }
            }
        ];

        return (
            <div className={`${styles.card} ${styles.detailedCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="LockSolid" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Resources with Unique Permissions ({permissions.length})</h2>
                </div>
                <div className={styles.cardBody}>
                    {permissions.length > 0 ? (
                        <DetailsList
                            items={permissions}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                        />
                    ) : (
                        <p>No resources have unique permissions.</p>
                    )}
                </div>
            </div>
        );
    }

    private renderAnonymousLinksTable(permissions: IPermissionIndicator[], siteUrl: string): React.ReactElement {
        const columns: IColumn[] = [
            { key: 'scope', name: 'Scope', fieldName: 'scope', minWidth: 80, maxWidth: 100 },
            { key: 'title', name: 'Resource Name', fieldName: 'title', minWidth: 200, maxWidth: 400, isResizable: true },
            { key: 'count', name: 'Anonymous Links Count', fieldName: 'anonymousLinkCount', minWidth: 150, maxWidth: 200 },
            { 
                key: 'action', 
                name: 'Action', 
                minWidth: 100, 
                onRender: (item: IPermissionIndicator) => {
                    const targetUrl = item.url || siteUrl;
                    return <Link href={targetUrl} target="_blank">View Resource</Link>;
                }
            }
        ];

        return (
            <div className={`${styles.card} ${styles.detailedCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Share" className={styles.cardHeaderIcon} style={{ color: '#ef4444' }} />
                    <h2 className={styles.cardHeaderTitle}>Resources with Anonymous Links ({permissions.length})</h2>
                </div>
                <div className={styles.cardBody}>
                    {permissions.length > 0 ? (
                        <DetailsList
                            items={permissions}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                        />
                    ) : (
                        <p>No anonymous sharing links found.</p>
                    )}
                </div>
            </div>
        );
    }

    private renderEmptyGroupsTable(groups: IGroupSummary[], siteUrl: string): React.ReactElement {
        const columns: IColumn[] = [
            { key: 'name', name: 'Group Name', fieldName: 'name', minWidth: 200, maxWidth: 300, isResizable: true },
            { key: 'id', name: 'Group ID', fieldName: 'id', minWidth: 100, maxWidth: 100 },
            { 
                key: 'action', 
                name: 'Action', 
                minWidth: 150, 
                onRender: (item: IGroupSummary) => (
                    <Link href={`${siteUrl}/_layouts/15/people.aspx?MembershipGroupId=${item.id}`} target="_blank">View Group</Link>
                )
            }
        ];

        return (
            <div className={`${styles.card} ${styles.detailedCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Delete" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Stale / Empty Groups ({groups.length})</h2>
                </div>
                <div className={styles.cardBody}>
                    {groups.length > 0 ? (
                        <DetailsList
                            items={groups}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                        />
                    ) : (
                        <p>No empty groups found.</p>
                    )}
                </div>
            </div>
        );
    }

// ═══════════════════════════════════
    // Sub-renders
    // ═══════════════════════════════════

    private renderMetricCard(
        icon: string,
        label: string,
        value: string,
        statusText: string,
        statusType: string,
        colorValue: boolean = false
    ): React.ReactElement {
        const statusClassMap: Record<string, string> = {
            Success: styles.metricStatusSuccess,
            Warning: styles.metricStatusWarning,
            Danger: styles.metricStatusDanger,
            Info: styles.metricStatusInfo,
            Neutral: styles.metricStatusNeutral,
        };
        return (
            <div className={styles.metricCard}>
                <div className={styles.metricIconWrap}>
                    <Icon iconName={icon} />
                </div>
                <div className={styles.metricBody}>
                    <p className={styles.metricLabel}>{label}</p>
                    <p className={`${styles.metricValue} ${colorValue ? (statusClassMap[statusType] || '') : ''}`}>{value}</p>
                    <p className={`${styles.metricStatus} ${statusClassMap[statusType] || styles.metricStatusNeutral}`}>
                        {statusText}
                    </p>
                </div>
            </div>
        );
    }

    // ── Risk Assessment Card ──

    private renderRiskAssessmentCard(result: IScanResult, isUnified: boolean = false): React.ReactElement {
        const { riskScore } = result;
        const ringColor = this.getRiskColor(riskScore.score);

        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Shield" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Assessment Summary</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={isUnified ? '' : styles.riskAssessmentBody}>
                        {/* Gauge panel */}
                        {!isUnified && (
                            <div className={styles.gaugePanel}>
                                <div
                                    className={styles.riskRing}
                                    style={{
                                        background: `conic-gradient(${ringColor} ${riskScore.score * 3.6}deg, #e5e7eb ${riskScore.score * 3.6}deg)`,
                                    }}
                                >
                                    <div className={styles.riskRingInner} />
                                    <div className={styles.riskRingContent}>
                                        <span className={styles.riskNumber} style={{ color: ringColor }}>
                                            {riskScore.score}
                                        </span>
                                        <span className={styles.riskSubLabel}>/ 100</span>
                                    </div>
                                </div>
                                <span className={styles.riskLevelText} style={{ color: ringColor }}>
                                    {riskScore.label} Risk
                                </span>
                            </div>
                        )}

                        {/* Summary panel */}
                        <div className={styles.riskSummaryPanel} style={isUnified ? { flex: '1 1 100%', display: 'flex', flexDirection: 'row', gap: '40px' } : {}}>
                            <div style={isUnified ? { flex: '1 1 50%' } : {}}>
                                {/* In unified we removed the <h3> because the card header is already Assessment Summary */}
                                {!isUnified && <h3>Assessment Summary</h3>}
                                <p style={isUnified ? { fontSize: '15px', lineHeight: '1.6' } : {}}>
                                    {this.getRiskSummaryText(riskScore.label, riskScore.score)}
                                </p>
                            </div>

                            {riskScore.scoreBreakdown && riskScore.scoreBreakdown.length > 0 && (
                                <div style={isUnified ? { flex: '1 1 50%' } : {}}>
                                    <h4 style={isUnified ? { marginTop: 0 } : {}}>Score Breakdown</h4>
                                    <table className={styles.dataTable} style={{ marginTop: '10px' }}>
                                        <thead>
                                            <tr>
                                                <th>Finding</th>
                                                <th style={{ textAlign: 'right' }}>Impact</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {riskScore.scoreBreakdown.map((b, idx) => (
                                                <tr key={idx}>
                                                    <td>{b.finding}</td>
                                                    <td style={{ fontWeight: 'bold', color: b.points > 0 ? '#d83b01' : 'inherit', textAlign: 'right' }}>
                                                        {b.points > 0 ? `+${b.points}` : b.points}
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    // ── Key Findings Card ──

    private renderKeyFindingsCard(result: IScanResult): React.ReactElement {
        const { users, groups, permissionIndicators } = result;
        const externalUsers = users.filter((u) => u.isPossibleExternal);
        const uniquePerms = permissionIndicators.filter((p) => p.hasUniquePermissions);
        const groupsWithExt = groups.filter(
            (g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0
        );

        const findings: { title: string; desc: string; type: string }[] = [];

        if (externalUsers.length === 0) {
            findings.push({
                title: 'No External Users Detected',
                desc: 'All users appear to be internal domain members.',
                type: 'success',
            });
        } else {
            findings.push({
                title: `${externalUsers.length} External User(s) Found`,
                desc: `Possible external users: ${externalUsers.slice(0, 3).map((u) => u.displayName).join(', ')}${externalUsers.length > 3 ? '...' : ''}`,
                type: 'warning',
            });
        }

        if (groupsWithExt.length > 0) {
            findings.push({
                title: `${groupsWithExt.length} Group(s) Contain External Users`,
                desc: groupsWithExt.map((g) => g.name).join(', '),
                type: 'warning',
            });
        } else {
            findings.push({
                title: 'Groups Are Internal Only',
                desc: 'No external users found in any SharePoint group.',
                type: 'success',
            });
        }

        if (uniquePerms.length > 0) {
            findings.push({
                title: `${uniquePerms.length} Unique Permission(s)`,
                desc: 'Some resources have broken inheritance.',
                type: 'info',
            });
        } else {
            findings.push({
                title: 'Consistent Permissions',
                desc: 'All items inherit permissions from the site.',
                type: 'success',
            });
        }

        if (result.errors.length > 0) {
            findings.push({
                title: `${result.errors.length} Scan Error(s)`,
                desc: result.errors[0].message,
                type: 'danger',
            });
        }

        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Lightbulb" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Key Findings</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.findingsList}>
                        {findings.map((f, idx) => this.renderFinding(f, idx))}
                    </div>
                </div>
            </div>
        );
    }

    private renderFinding(
        finding: { title: string; desc: string; type: string },
        idx: number
    ): React.ReactElement {
        const typeMap: Record<string, { row: string; icon: string; iconSymbol: string }> = {
            success: { row: styles.findingSuccess, icon: styles.findingIconSuccess, iconSymbol: '✓' },
            warning: { row: styles.findingWarning, icon: styles.findingIconWarning, iconSymbol: '!' },
            info: { row: styles.findingInfo, icon: styles.findingIconInfo, iconSymbol: 'i' },
            danger: { row: styles.findingDanger, icon: styles.findingIconDanger, iconSymbol: '✕' },
        };
        const t = typeMap[finding.type] || typeMap.info;
        return (
            <div className={`${styles.findingItem} ${t.row}`} key={idx}>
                <div className={`${styles.findingIcon} ${t.icon}`}>{t.iconSymbol}</div>
                <div className={styles.findingBody}>
                    <p className={styles.findingTitle}>{finding.title}</p>
                    <p className={styles.findingDescription}>{finding.desc}</p>
                </div>
            </div>
        );
    }

    // ── Groups Card ──

    private renderGroupsCard(groups: IGroupSummary[]): React.ReactElement {
        const groupsWithExternal = groups.filter(
            (g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0
        );
        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Group" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>SharePoint Groups</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.cardStatsRow}>
                        <div>
                            <p className={styles.statValue}>{groups.length}</p>
                            <p className={styles.statLabel}>Total Groups</p>
                        </div>
                        <div>
                            <p className={`${styles.statValue} ${groupsWithExternal.length > 0 ? styles.warningText : styles.successText}`}>
                                {groupsWithExternal.length}
                            </p>
                            <p className={styles.statLabel}>With External</p>
                        </div>
                    </div>

                    {groupsWithExternal.length > 0 && (
                        <div className={styles.inlineWarningBanner}>
                            <Icon iconName="Warning" />
                            <span>{groupsWithExternal.length} group(s) contain possible external users</span>
                        </div>
                    )}

                    <div className={styles.tableWrap}>
                        <table className={styles.dataTable}>
                            <thead>
                                <tr>
                                    <th>Group Name</th>
                                    <th>Members</th>
                                    <th>External</th>
                                </tr>
                            </thead>
                            <tbody>
                                {groups.slice(0, 8).map((g, idx) => (
                                    <tr key={idx}>
                                        <td>{g.name}</td>
                                        <td>{g.userCount ?? '—'}</td>
                                        <td className={
                                            (g.possibleExternalUserCount ?? 0) > 0
                                                ? styles.statusReview
                                                : styles.statusOk
                                        }>
                                            {g.possibleExternalUserCount ?? 0}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        );
    }

    // ── Permissions Card ──

    private renderPermissionsCard(permissions: IPermissionIndicator[]): React.ReactElement {
        const uniqueCount = permissions.filter((p) => p.hasUniquePermissions).length;
        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Lock" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Permission Inheritance</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.cardStatsRow}>
                        <div>
                            <p className={styles.statValue}>{permissions.length}</p>
                            <p className={styles.statLabel}>Items Reviewed</p>
                        </div>
                        <div>
                            <p className={`${styles.statValue} ${uniqueCount > 0 ? styles.warningText : styles.successText}`}>
                                {uniqueCount}
                            </p>
                            <p className={styles.statLabel}>Unique Permissions</p>
                        </div>
                    </div>

                    <div className={styles.tableWrap}>
                        <table className={styles.dataTable}>
                            <thead>
                                <tr>
                                    <th>Resource</th>
                                    <th>Scope</th>
                                    <th>Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                {permissions.slice(0, 8).map((p, idx) => (
                                    <tr key={idx}>
                                        <td>{p.title}</td>
                                        <td>{p.scope}</td>
                                        <td className={
                                            p.status === 'Ok'
                                                ? styles.statusOk
                                                : p.status === 'Review'
                                                    ? styles.statusReview
                                                    : styles.statusEmpty
                                        }>
                                            {p.status}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        );
    }

    // ── External User Indicators Card ──

    private renderExternalIndicatorsCard(externalUsers: IUserRiskInfo[]): React.ReactElement {
        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="ContactInfo" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>External User Indicators</h2>
                </div>
                <div className={styles.cardBody}>
                    {externalUsers.length === 0 ? (
                        <div className={styles.successPanelLarge}>
                            <div className={styles.successCircleIcon}>✓</div>
                            <p className={styles.successPanelTitle}>No External Users Detected</p>
                            <p className={styles.successPanelText}>
                                All site users appear to belong to internal domains.
                                No external sharing indicators were found.
                            </p>
                        </div>
                    ) : (
                        <div className={styles.dangerPanelLarge}>
                            <div className={styles.dangerCircleIcon}>!</div>
                            <p className={styles.dangerPanelTitle}>
                                {externalUsers.length} External User(s) Detected
                            </p>
                            <p className={styles.dangerPanelText}>
                                {externalUsers.slice(0, 4).map((u) => u.displayName).join(', ')}
                                {externalUsers.length > 4 ? ` and ${externalUsers.length - 4} more` : ''}
                            </p>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    // ── Recommended Actions ──

    private renderActionsCard(actions: string[]): React.ReactElement {
        const actionIcons = [
            'Shield', 'People', 'Lock', 'Settings', 'SkypeCheck', 'DocumentSearch',
        ];
        return (
            <div className={`${styles.card} ${styles.recommendedActionsCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Rocket" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Recommended Actions</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.actionGrid}>
                        {actions.map((action, idx) => (
                            <div className={styles.actionTile} key={idx}>
                                <div className={styles.actionIconWrap}>
                                    <Icon iconName={actionIcons[idx % actionIcons.length]} />
                                </div>
                                <span className={styles.actionText}>{action}</span>
                            </div>
                        ))}
                    </div>
                </div>
            </div>
        );
    }

    // ═══════════════════════════════════
    // Helpers
    // ═══════════════════════════════════

    private getRiskColor(score: number): string {
        if (score < 30) return '#10b981'; // Green
        if (score < 70) return '#f59e0b'; // Orange
        return '#ef4444'; // Red
    }

    private getRiskStatusClass(score: number): string {
        if (score < 30) return 'Success';
        if (score < 70) return 'Warning';
        return 'Danger';
    }

    private getRiskSummaryText(label: RiskLabel, score: number): React.ReactNode {
        const color = this.getRiskColor(score);
        const scoreSpan = <span style={{ color, fontWeight: 'bold' }}>{score}/100</span>;

        if (label === 'Low') {
            return <>This site has a low risk score of {scoreSpan}. No significant external sharing concerns were identified. The site permissions appear well-configured and all users belong to expected internal domains.</>;
        }
        if (label === 'Medium') {
            return <>This site has a moderate risk score of {scoreSpan}. Some external sharing signals were detected that warrant review. Consider checking the contributing factors below and following the recommended actions.</>;
        }
        if (label === 'High') {
            return <>This site has an elevated risk score of {scoreSpan}. Multiple external sharing concerns were detected. Immediate review of external users and permission configurations is recommended.</>;
        }
        if (label === 'Review Recommended') {
            return <>This site scored {scoreSpan} and requires further review. Some aspects of the permissions or user configuration could not be fully verified.</>;
        }
        return <>Risk assessment could not be fully completed. Score: {scoreSpan}.</>;
    }

    private formatDate(dateStr: string): string {
        try {
            const d = new Date(dateStr);
            return d.toLocaleString(undefined, {
                year: 'numeric',
                month: 'short',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
            });
        } catch {
            return dateStr;
        }
    }
}
