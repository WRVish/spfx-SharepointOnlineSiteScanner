export type RiskLabel = 'Low' | 'Medium' | 'High' | 'Review Recommended' | 'Unknown';

export interface IRiskScore {
    score: number;
    label: RiskLabel;
    reasons: string[];
    scoreBreakdown: { finding: string; points: number }[];
    recommendedActions: string[];
}
