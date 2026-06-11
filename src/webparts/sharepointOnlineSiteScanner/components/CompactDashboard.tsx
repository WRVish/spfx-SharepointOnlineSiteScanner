import * as React from 'react';
import styles from './CompactDashboard.module.scss';
import { IScanResult } from '../models/IScanResult';

export interface ICompactDashboardProps {
    scanResult: IScanResult;
    onExportCsv: () => void;
}

export class CompactDashboard extends React.Component<ICompactDashboardProps> {
    public render(): React.ReactElement {
        const { scanResult } = this.props;
        const { riskScore, users, permissionIndicators, groups } = scanResult;

        const externalUsers = users.filter(u => u.isPossibleExternal);
        const uniquePerms = permissionIndicators.filter(p => p.hasUniquePermissions);
        const librariesReviewed = permissionIndicators.filter(p => p.scope === 'Library').length;
        const groupsReviewed = groups.length;
        
        const riskColor = riskScore.score < 30 ? '#10b981' : riskScore.score < 70 ? '#f59e0b' : '#ef4444';
        const riskPillClass = riskScore.score < 30 ? styles.pillSuccess : riskScore.score < 70 ? styles.pillWarning : styles.pillDanger;
        
        // Find stale libraries (>2 years)
        const staleLibs = permissionIndicators.filter(p => {
            if (!p.lastModified || p.scope !== 'Library') return false;
            const modifiedDate = new Date(p.lastModified);
            const twoYearsAgo = new Date();
            twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);
            return modifiedDate < twoYearsAgo;
        });

        // Find Copilot Overshared Resources
        const overshared = permissionIndicators.filter(p => p.isOvershared);
        
        // Find Overshared Groups
        const oversharedGroups = groups.filter(g => g.isOvershared);

        return (
            <div className={styles.compactRoot}>
                <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '16px', gap: '10px' }}>
                    <button style={{ padding: '6px 12px', cursor: 'pointer', borderRadius: '4px', border: '1px solid #d1d5db', background: 'white' }} onClick={this.props.onExportCsv}>
                        Export CSV
                    </button>
                </div>
                {/* TOP ROW: Metric Cards */}
                <div className={styles.topRow}>
                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Risk Score</span>
                            <span className={styles.compactValue} style={{ color: riskColor }}>
                                {riskScore.score} <span style={{ fontSize: '14px', color: '#6b7280' }}>/ 100</span>
                            </span>
                        </div>
                        <span className={`${styles.pill} ${riskPillClass}`}>
                            {riskScore.label}
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>External Users</span>
                            <span className={styles.compactValue}>{externalUsers.length}</span>
                        </div>
                        <span className={`${styles.pill} ${externalUsers.length > 0 ? styles.pillDanger : styles.pillSuccess}`}>
                            {externalUsers.length > 0 ? 'Review' : 'OK'}
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Groups Reviewed</span>
                            <span className={styles.compactValue}>{groupsReviewed}</span>
                        </div>
                        <span className={`${styles.pill} ${styles.pillSuccess}`}>
                            Scanned
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Unique Perms</span>
                            <span className={styles.compactValue}>{uniquePerms.length}</span>
                        </div>
                        <span className={`${styles.pill} ${uniquePerms.length > 0 ? styles.pillWarning : styles.pillSuccess}`}>
                            Isolated
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Libraries Reviewed</span>
                            <span className={styles.compactValue}>{librariesReviewed}</span>
                        </div>
                        <span className={`${styles.pill} ${styles.pillSuccess}`}>
                            Scanned
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Stale Libraries</span>
                            <span className={styles.compactValue}>{staleLibs.length}</span>
                        </div>
                        <span className={`${styles.pill} ${staleLibs.length > 0 ? styles.pillWarning : styles.pillSuccess}`}>
                            &gt; 2 Years
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Overshared Grp</span>
                            <span className={styles.compactValue}>{oversharedGroups.length}</span>
                        </div>
                        <span className={`${styles.pill} ${oversharedGroups.length > 0 ? styles.pillDanger : styles.pillSuccess}`}>
                            Everyone
                        </span>
                    </div>

                    <div className={styles.compactCard}>
                        <div className={styles.compactCardInfo}>
                            <span className={styles.compactLabel}>Total Users</span>
                            <span className={styles.compactValue}>{users.length}</span>
                        </div>
                        <span className={styles.pill}>
                            Indexed
                        </span>
                    </div>
                </div>

                {/* MIDDLE ROW: Score Breakdown & Ext Users */}
                <div className={styles.middleRow}>
                    <div className={styles.middlePanel}>
                        <h2 className={styles.sectionTitle}>Score Breakdown</h2>
                        <table className={styles.dataTable}>
                            <thead>
                                <tr>
                                    <th>Finding</th>
                                    <th style={{ textAlign: 'right' }}>Impact</th>
                                </tr>
                            </thead>
                            <tbody>
                                {riskScore.scoreBreakdown?.map((b, idx) => (
                                    <tr key={idx}>
                                        <td>{b.finding}</td>
                                        <td style={{ textAlign: 'right', fontWeight: 600, color: b.points > 0 ? '#ef4444' : '#10b981' }}>
                                            {b.points > 0 ? `+${b.points}` : b.points}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className={styles.middlePanel}>
                        <h2 className={styles.sectionTitle}>External Users Snapshot</h2>
                        {externalUsers.length > 0 ? (
                            <table className={styles.dataTable}>
                                <thead>
                                    <tr>
                                        <th>Name</th>
                                        <th>Email</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {externalUsers.map((u, idx) => (
                                        <tr key={idx}>
                                            <td style={{ fontWeight: 500 }}>{u.displayName}</td>
                                            <td style={{ color: '#6b7280' }}>{u.email || u.loginName}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#6b7280', fontSize: '13px' }}>No external users detected.</p>
                        )}
                    </div>
                </div>

                <div className={styles.middleRow} style={{ marginTop: '24px' }}>
                    <div className={styles.middlePanel}>
                        <h2 className={styles.sectionTitle}>Groups with External Users</h2>
                        {groups.filter(g => (g.possibleExternalUserCount ?? 0) > 0).length > 0 ? (
                            <table className={styles.dataTable}>
                                <thead>
                                    <tr>
                                        <th>Group Name</th>
                                        <th>Total Users</th>
                                        <th>External Count</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {groups.filter(g => (g.possibleExternalUserCount ?? 0) > 0).map((g, idx) => (
                                        <tr key={idx}>
                                            <td style={{ fontWeight: 500 }}>{g.name}</td>
                                            <td>{g.userCount}</td>
                                            <td><span className={`${styles.pill} ${styles.pillWarning}`}>{g.possibleExternalUserCount}</span></td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#6b7280', fontSize: '13px' }}>No exposed groups found.</p>
                        )}
                    </div>
                </div>

                {/* BOTTOM ROW: Copilot Readiness and Unique Permissions */}
                <div className={styles.middleRow}>
                    <div className={styles.middlePanel}>
                        <h2 className={styles.sectionTitle}>Copilot Readiness: Overshared Resources</h2>
                        {overshared.length > 0 ? (
                            <table className={styles.dataTable}>
                                <thead>
                                    <tr>
                                        <th>Resource Name</th>
                                        <th>Scope</th>
                                        <th>Anonymous Links</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {overshared.map((p, idx) => (
                                        <tr key={idx}>
                                            <td style={{ fontWeight: 500, color: '#ef4444' }}>{p.title}</td>
                                            <td>{p.scope}</td>
                                            <td>{p.anonymousLinkCount || 0}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#047857', fontSize: '13px', fontWeight: 500 }}>
                                ✓ No resources are overshared with "Everyone" or "All Company".
                            </p>
                        )}
                    </div>
                    
                    <div className={styles.middlePanel}>
                        <h2 className={styles.sectionTitle}>All Unique Permissions</h2>
                        {uniquePerms.length > 0 ? (
                            <table className={styles.dataTable}>
                                <thead>
                                    <tr>
                                        <th>Resource Name</th>
                                        <th>Scope</th>
                                        <th>Overshared?</th>
                                        <th>Anon Links</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {uniquePerms.map((p, idx) => (
                                        <tr key={idx}>
                                            <td style={{ fontWeight: 500 }}>{p.title}</td>
                                            <td>{p.scope}</td>
                                            <td><span className={`${styles.pill} ${p.isOvershared ? styles.pillDanger : styles.pillSuccess}`}>{p.isOvershared ? 'Yes' : 'No'}</span></td>
                                            <td style={{color: p.anonymousLinkCount && p.anonymousLinkCount > 0 ? '#ef4444' : 'inherit'}}>{p.anonymousLinkCount || 0}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#6b7280', fontSize: '13px' }}>No unique permissions found.</p>
                        )}
                    </div>
                </div>

            </div>
        );
    }
}
