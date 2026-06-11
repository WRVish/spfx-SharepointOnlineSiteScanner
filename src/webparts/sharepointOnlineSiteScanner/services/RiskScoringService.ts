import { IScanResult } from '../models/IScanResult';
import { IRiskScore, RiskLabel } from '../models/IRiskScore';
import {
    RISK_WEIGHTS,
    BUCKET_CAPS,
    RISK_THRESHOLDS,
    MAX_RISK_SCORE,
    MIN_RISK_SCORE,
    EXTERNAL_USER_THRESHOLDS,
} from '../constants/RiskConstants';
import { LoggerService } from './LoggerService';

const SOURCE = 'RiskScoringService';

export class RiskScoringService {
    /**
     * Calculates a deterministic risk score based on scan results.
     * The scoring logic is transparent and easy to explain.
     *
     * @param scanPartial - Partial scan result data to score
     * @returns Calculated risk score with reasons and recommended actions
     */
    public calculateRiskScore(scanPartial: Partial<IScanResult>, customCaps?: any): IRiskScore {
        LoggerService.info(SOURCE, 'Calculating risk score with bucketed model...');
        
        // Use custom caps from Web Part properties if provided and valid (sum to 100), otherwise fallback to constants
        const caps = customCaps || BUCKET_CAPS;
        
        let bucketOversharing = 0;
        let bucketExternalUser = 0;
        let bucketFileSharing = 0;
        let bucketGovernance = 0;

        const reasons: string[] = [];
        const recommendedActions: string[] = [];
        const scoreBreakdown: { finding: string; points: number }[] = [];

        // ==========================================
        // BUCKET 1: OVERSHARING RISK (Max 40)
        // ==========================================
        let isCopilotOvershared = false;

        if (scanPartial.permissionIndicators) {
            const oversharedItems = scanPartial.permissionIndicators.filter(p => p.isOvershared);
            if (oversharedItems.length > 0) isCopilotOvershared = true;
        }

        if (!isCopilotOvershared && scanPartial.users) {
            const hasEEEU = scanPartial.users.some(u => 
                u.displayName.toLowerCase().includes('everyone') || 
                u.displayName.toLowerCase().includes('all company')
            );
            if (hasEEEU) isCopilotOvershared = true;
        }

        if (isCopilotOvershared) {
            bucketOversharing += RISK_WEIGHTS.SITE_EXPOSED_TO_EVERYONE;
            reasons.push('Oversharing Detected: Site is exposed to Everyone/All Company');
            recommendedActions.push('CRITICAL: Remove "Everyone" claims to ensure Copilot readiness and prevent data leaks');
        }

        if (scanPartial.groups) {
            let oversharedGroups = 0;
            for (const group of scanPartial.groups) {
                if (group.isOvershared) oversharedGroups++;
            }
            if (oversharedGroups > 0) {
                bucketOversharing += oversharedGroups * RISK_WEIGHTS.GROUP_EXPOSED_TO_EVERYONE;
                reasons.push(`${oversharedGroups} SharePoint group(s) contain "Everyone" (Major AI Data Risk)`);
                recommendedActions.push('Remove "Everyone" or "All Company" claims from SharePoint groups to prevent oversharing');
            }
        }

        bucketOversharing = Math.min(bucketOversharing, caps.OVERSHARING_RISK);
        if (bucketOversharing > 0) {
            scoreBreakdown.push({ finding: 'Copilot / Oversharing Risk', points: bucketOversharing });
        }

        // ==========================================
        // BUCKET 2: EXTERNAL USER RISK (Max 30)
        // ==========================================
        if (scanPartial.users) {
            const externalUsers = scanPartial.users.filter((u) => u.isPossibleExternal);

            if (externalUsers.length > 0) {
                bucketExternalUser += RISK_WEIGHTS.EXTERNAL_USER_FOUND;
                reasons.push(`${externalUsers.length} possible external user(s) detected in the site`);
                recommendedActions.push('Review external users in SharePoint site permissions to confirm they still need access');

                if (externalUsers.length > EXTERNAL_USER_THRESHOLDS.MANY) {
                    bucketExternalUser += RISK_WEIGHTS.MANY_EXTERNAL_USERS;
                    reasons.push(`More than ${EXTERNAL_USER_THRESHOLDS.MANY} external users found — elevated risk`);
                }
            }
        }

        if (scanPartial.groups) {
            const groupsWithExternal = scanPartial.groups.filter(g => g.possibleExternalUserCount && g.possibleExternalUserCount > 0);

            if (groupsWithExternal.length > 0) {
                const ownerGroups = groupsWithExternal.filter(g => g.name.toLowerCase().includes('owner'));
                if (ownerGroups.length > 0) {
                    bucketExternalUser += RISK_WEIGHTS.EXTERNAL_USERS_IN_OWNERS;
                    reasons.push('External users found in Owner group(s)');
                    recommendedActions.push('Review Owner group membership — external users in Owner groups have elevated permissions');
                }

                const memberGroups = groupsWithExternal.filter(g => g.name.toLowerCase().includes('member'));
                if (memberGroups.length > 0) {
                    bucketExternalUser += RISK_WEIGHTS.EXTERNAL_USERS_IN_MEMBERS;
                    reasons.push('External users found in Member group(s)');
                    recommendedActions.push('Confirm external users in Member groups should have edit-level access');
                }
            }
        }

        bucketExternalUser = Math.min(bucketExternalUser, caps.EXTERNAL_USER_RISK);
        if (bucketExternalUser > 0) {
            scoreBreakdown.push({ finding: 'External User Risk', points: bucketExternalUser });
        }

        // ==========================================
        // BUCKET 3: FILE SHARING RISK (Max 20)
        // ==========================================
        if (scanPartial.permissionIndicators) {
            const siteIndicators = scanPartial.permissionIndicators.filter(p => p.scope === 'Site');
            const libraryIndicators = scanPartial.permissionIndicators.filter(p => p.scope === 'Library');

            const siteHasUnique = siteIndicators.some(p => p.hasUniquePermissions === true);
            const librariesWithUnique = libraryIndicators.filter(p => p.hasUniquePermissions === true);

            if (siteHasUnique || librariesWithUnique.length > 0) {
                bucketFileSharing += RISK_WEIGHTS.UNIQUE_PERMISSIONS;
                reasons.push('Site or libraries have unique (broken) permission inheritance');
                recommendedActions.push('Review permission inheritance — unique permissions increase governance complexity');
            }

            const anonLinkCount = scanPartial.permissionIndicators.reduce((sum, p) => sum + (p.anonymousLinkCount || 0), 0);
            if (anonLinkCount > 0) {
                bucketFileSharing += anonLinkCount * RISK_WEIGHTS.ANONYMOUS_LINK;
                reasons.push(`Found ${anonLinkCount} active Anonymous (Anyone) link(s)`);
                recommendedActions.push('Review and revoke anonymous sharing links to secure sensitive documents');
            }
        }

        bucketFileSharing = Math.min(bucketFileSharing, caps.FILE_SHARING_RISK);
        if (bucketFileSharing > 0) {
            scoreBreakdown.push({ finding: 'File Sharing & Links Risk', points: bucketFileSharing });
        }

        // ==========================================
        // BUCKET 4: GOVERNANCE HYGIENE (Max 10)
        // ==========================================
        if (scanPartial.errors && scanPartial.errors.length > 0) {
            bucketGovernance += RISK_WEIGHTS.DATA_RETRIEVAL_FAILURE;
            reasons.push(`${scanPartial.errors.length} error(s) occurred during scan — some data may be incomplete`);
            recommendedActions.push('Review scan errors and ensure you have sufficient permissions to access site data');
        }

        if (scanPartial.groups) {
            const emptyGroups = scanPartial.groups.filter(g => g.isEmpty === true);
            if (emptyGroups.length > 0) {
                bucketGovernance += emptyGroups.length * RISK_WEIGHTS.EMPTY_GROUP;
                reasons.push(`${emptyGroups.length} empty group(s) found — may indicate stale permissions`);
                recommendedActions.push('Review empty SharePoint groups and remove if no longer needed');
            }
        }

        if (scanPartial.permissionIndicators) {
            const staleLibs = scanPartial.permissionIndicators.filter(p => {
                if (!p.lastModified) return false;
                const modifiedDate = new Date(p.lastModified);
                const twoYearsAgo = new Date();
                twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);
                return modifiedDate < twoYearsAgo;
            });
            if (staleLibs.length > 0) {
                bucketGovernance += staleLibs.length * RISK_WEIGHTS.STALE_LIBRARY;
                reasons.push(`${staleLibs.length} document libraries appear stale (not modified in >2 years)`);
                recommendedActions.push('Review stale libraries and archive them to improve Copilot accuracy');
            }
        }

        bucketGovernance = Math.min(bucketGovernance, caps.GOVERNANCE_HYGIENE);
        if (bucketGovernance > 0) {
            scoreBreakdown.push({ finding: 'Governance Hygiene Risk', points: bucketGovernance });
        }

        // ==========================================
        // FINAL SCORE CALCULATION
        // ==========================================
        let score = bucketOversharing + bucketExternalUser + bucketFileSharing + bucketGovernance;

        // Determine label
        const label = this.getLabel(score, scanPartial);

        // Add general recommended actions
        if (recommendedActions.length === 0) {
            recommendedActions.push('No significant risk indicators found — continue routine governance reviews');
        }

        recommendedActions.push('Validate sharing settings with your Microsoft 365 admin');
        recommendedActions.push('Confirm whether external access is still needed before migration or Copilot rollout');

        const result: IRiskScore = {
            score,
            label,
            reasons: reasons.length > 0 ? reasons : ['No significant risk indicators detected'],
            scoreBreakdown,
            recommendedActions,
        };

        LoggerService.info(SOURCE, `Risk score calculated: ${score} (${label})`);
        return result;
    }

    /**
     * Maps a numeric score to a risk label.
     */
    private getLabel(score: number, scanPartial: Partial<IScanResult>): RiskLabel {
        // If there were significant errors, show "Unknown"
        if (
            scanPartial.errors &&
            scanPartial.errors.length > 2 &&
            (!scanPartial.users || scanPartial.users.length === 0)
        ) {
            return 'Unknown';
        }

        if (score <= RISK_THRESHOLDS.LOW_MAX) return 'Low';
        if (score <= RISK_THRESHOLDS.MEDIUM_MAX) return 'Medium';
        if (score <= RISK_THRESHOLDS.HIGH_MAX) return 'High';
        return 'Review Recommended';
    }
}
