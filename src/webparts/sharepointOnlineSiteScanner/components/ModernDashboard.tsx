import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './ModernDashboard.module.scss';
import { IScanResult } from '../models/IScanResult';
import { RiskLabel } from '../models/IRiskScore';

export interface IModernDashboardProps {
    scanResult: IScanResult;
    onRescan: () => void;
    onExportCsv: () => void;
}

export class ModernDashboard extends React.Component<IModernDashboardProps> {

    private getRiskColor(score: number): string {
        if (score < 30) return '#10b981'; // Green
        if (score < 70) return '#f59e0b'; // Orange
        return '#ef4444'; // Red
    }

    public render(): React.ReactElement {
        const { scanResult, onRescan } = this.props;
        const { riskScore, users, permissionIndicators, siteSummary, groups } = scanResult;

        const externalUsers = users.filter(u => u.isPossibleExternal);
        const groupsWithExt = groups.filter(g => g.possibleExternalUserCount && g.possibleExternalUserCount > 0);
        const uniquePermCount = permissionIndicators.filter(p => p.hasUniquePermissions).length;
        const librariesReviewed = permissionIndicators.filter(p => p.scope === 'Library').length;
        const groupsReviewed = groups.length;

        const ringColor = this.getRiskColor(riskScore.score);
        const deg = riskScore.score * 3.6;

        return (
            <div className={styles.modernRoot}>
                <div className={styles.topHeader}>
                    <h1>Site Security Pulse</h1>
                    <div style={{ display: 'flex', gap: '10px' }}>
                        <button className={`${styles.actionButton}`} onClick={this.props.onExportCsv}>
                            <Icon iconName="DownloadDocument" /> Export CSV
                        </button>
                        <button className={`${styles.actionButton} ${styles.primaryButton}`} onClick={onRescan}>
                            <Icon iconName="Refresh" /> Scan Now
                        </button>
                    </div>
                </div>

                <div className={styles.bentoGrid}>
                    
                    {/* Risk Gauge */}
                    <div className={`${styles.bentoCard} ${styles.riskCard}`}>
                        <div className={styles.gaugeContainer}>
                            <div className={styles.gaugeRing} style={{ background: `conic-gradient(${ringColor} ${deg}deg, #e5e7eb ${deg}deg)` }}></div>
                            <div className={styles.gaugeInner}>
                                <span className={styles.gaugeScore} style={{ fontSize: '24px' }}>{riskScore.score} <span style={{ fontSize: '14px', color: '#6b7280' }}>/ 100</span></span>
                            </div>
                        </div>
                        <h2 className={styles.riskLabel} style={{ color: ringColor }}>{riskScore.label} Risk</h2>
                        <p className={styles.metricLabel}>Overall Assessment</p>
                    </div>

                    {/* Key Findings Summary */}
                    <div className={`${styles.bentoCard} ${styles.summaryCard}`}>
                        <div style={{ display: 'flex', gap: '40px' }}>
                            <div style={{ flex: 1 }}>
                                <h2 style={{marginTop: 0, color: '#111827', fontSize: '18px'}}>Critical Findings</h2>
                                <ul style={{ paddingLeft: '20px', color: '#4b5563', lineHeight: '1.6', margin: 0 }}>
                                    {riskScore.reasons.map((reason, idx) => (
                                        <li key={idx}><strong>{reason}</strong></li>
                                    ))}
                                </ul>
                                <div style={{ marginTop: '16px' }}>
                                    <span className={styles.metricLabel}>Last Scanned: {new Date(siteSummary.scannedAt).toLocaleString()}</span>
                                </div>
                            </div>
                            
                            {riskScore.scoreBreakdown && riskScore.scoreBreakdown.length > 0 && (
                                <div style={{ flex: 1 }}>
                                    <h2 style={{marginTop: 0, color: '#111827', fontSize: '18px'}}>Score Breakdown</h2>
                                    <table className={styles.dataTable} style={{ marginTop: '4px' }}>
                                        <thead>
                                            <tr>
                                                <th>Finding</th>
                                                <th style={{ textAlign: 'right' }}>Impact</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {riskScore.scoreBreakdown.map((b, idx) => (
                                                <tr key={idx}>
                                                    <td style={{ color: '#4b5563', fontSize: '13px' }}>{b.finding}</td>
                                                    <td style={{ textAlign: 'right', fontWeight: 600, color: b.points > 0 ? '#ef4444' : '#10b981' }}>
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

                    {/* Metrics Row */}
                    <div className={`${styles.bentoCard} ${styles.metricCard}`}>
                        <span className={styles.metricLabel}>External Users</span>
                        <span className={styles.metricValue}>{externalUsers.length}</span>
                        <span className={`${styles.pill} ${externalUsers.length > 0 ? styles.pillDanger : styles.pillSuccess}`}>
                            {externalUsers.length > 0 ? 'Review Needed' : 'Secure'}
                        </span>
                    </div>

                    <div className={`${styles.bentoCard} ${styles.metricCard}`}>
                        <span className={styles.metricLabel}>Groups Reviewed</span>
                        <span className={styles.metricValue}>{groupsReviewed}</span>
                        <span className={`${styles.pill} ${groupsWithExt.length > 0 ? styles.pillWarning : styles.pillSuccess}`}>
                            {groupsWithExt.length > 0 ? `${groupsWithExt.length} Exposed` : 'Secure'}
                        </span>
                    </div>

                    <div className={`${styles.bentoCard} ${styles.metricCard}`}>
                        <span className={styles.metricLabel}>Unique Permissions</span>
                        <span className={styles.metricValue}>{uniquePermCount}</span>
                        <span className={`${styles.pill} ${uniquePermCount > 0 ? styles.pillWarning : styles.pillSuccess}`}>
                            {uniquePermCount > 0 ? 'Items Isolated' : 'Inheriting'}
                        </span>
                    </div>

                    <div className={`${styles.bentoCard} ${styles.metricCard}`}>
                        <span className={styles.metricLabel}>Libraries Reviewed</span>
                        <span className={styles.metricValue}>{librariesReviewed}</span>
                        <span className={`${styles.pill} ${styles.pillSuccess}`}>
                            Scanned
                        </span>
                    </div>

                    {/* Simplified Tables for Modern UI */}
                    <div className={`${styles.bentoCard}`} style={{ gridColumn: 'span 12' }}>
                        <h2 style={{marginTop: 0, color: '#111827', fontSize: '18px'}}>External Users Snapshot</h2>
                        {externalUsers.length > 0 ? (
                            <table className={styles.dataTable}>
                                <thead>
                                    <tr>
                                        <th>Name</th>
                                        <th>Email</th>
                                        <th>Status</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {externalUsers.map((u, i) => (
                                        <tr key={i}>
                                            <td style={{fontWeight: 500}}>{u.displayName}</td>
                                            <td style={{color: '#6b7280'}}>{u.email || u.loginName}</td>
                                            <td><span className={`${styles.pill} ${styles.pillDanger}`}>External</span></td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#6b7280' }}>No external users found. The perimeter is secure.</p>
                        )}
                    </div>

                    <div className={`${styles.bentoCard}`} style={{ gridColumn: 'span 12' }}>
                        <h2 style={{marginTop: 0, color: '#111827', fontSize: '18px'}}>Groups with External Users</h2>
                        {groupsWithExt.length > 0 ? (
                            <table className={styles.dataTable}>
                                <thead>
                                    <tr>
                                        <th>Group Name</th>
                                        <th>Total Users</th>
                                        <th>External Count</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {groupsWithExt.map((g, i) => (
                                        <tr key={i}>
                                            <td style={{fontWeight: 500}}>{g.name}</td>
                                            <td>{g.userCount}</td>
                                            <td><span className={`${styles.pill} ${styles.pillWarning}`}>{g.possibleExternalUserCount}</span></td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#6b7280' }}>No exposed groups found.</p>
                        )}
                    </div>

                    <div className={`${styles.bentoCard}`} style={{ gridColumn: 'span 12' }}>
                        <h2 style={{marginTop: 0, color: '#111827', fontSize: '18px'}}>Resources with Unique Permissions</h2>
                        {permissionIndicators.filter(p => p.hasUniquePermissions).length > 0 ? (
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
                                    {permissionIndicators.filter(p => p.hasUniquePermissions).map((p, i) => (
                                        <tr key={i}>
                                            <td style={{fontWeight: 500}}>{p.title}</td>
                                            <td>{p.scope}</td>
                                            <td><span className={`${styles.pill} ${p.isOvershared ? styles.pillDanger : styles.pillSuccess}`}>{p.isOvershared ? 'Yes' : 'No'}</span></td>
                                            <td style={{color: p.anonymousLinkCount && p.anonymousLinkCount > 0 ? '#ef4444' : 'inherit'}}>{p.anonymousLinkCount || 0}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        ) : (
                            <p style={{ color: '#6b7280' }}>No unique permissions found.</p>
                        )}
                    </div>

                </div>
            </div>
        );
    }
}
