/**
 * Risk scoring constants for the External Sharing Risk Scanner.
 * All thresholds and weights are transparent and deterministic.
 */

/** 
 * MAXIMUM BUCKET CAPS 
 * The total score is the sum of these 4 buckets. 
 * They MUST sum exactly to 100 to ensure a perfect 0-100 percentage scale.
 */
export const BUCKET_CAPS = {
    OVERSHARING_RISK: 40,   // Broad exposure (EEEU, All Company) - Highest priority for Copilot/AI
    EXTERNAL_USER_RISK: 30, // Guest access and external permissions
    FILE_SHARING_RISK: 20,  // Anonymous links and broken inheritance
    GOVERNANCE_HYGIENE: 10  // Stale data and empty groups
};

/** Internal point values assigned to specific findings before hitting the bucket cap */
export const RISK_WEIGHTS = {
    // Bucket 1: Oversharing
    SITE_EXPOSED_TO_EVERYONE: 40,
    GROUP_EXPOSED_TO_EVERYONE: 20,

    // Bucket 2: External Users
    EXTERNAL_USER_FOUND: 10,
    MANY_EXTERNAL_USERS: 10,
    EXTERNAL_USERS_IN_OWNERS: 20,
    EXTERNAL_USERS_IN_MEMBERS: 10,

    // Bucket 3: File Sharing
    ANONYMOUS_LINK: 10,
    UNIQUE_PERMISSIONS: 10,

    // Bucket 4: Governance Hygiene
    STALE_LIBRARY: 5,
    EMPTY_GROUP: 5,
    DATA_RETRIEVAL_FAILURE: 10,
};

/** Risk label thresholds based on score (0-100) */
export const RISK_THRESHOLDS = {
    LOW_MAX: 25,
    MEDIUM_MAX: 50,
    HIGH_MAX: 75,
    // Anything above HIGH_MAX = "Review Recommended"
};

/** Maximum score cap */
export const MAX_RISK_SCORE = 100;

/** Minimum score floor */
export const MIN_RISK_SCORE = 0;

/** Maximum number of libraries to scan for permissions */
export const MAX_LIBRARIES_TO_SCAN = 20;

/** External user count thresholds */
export const EXTERNAL_USER_THRESHOLDS = {
    /** Trigger "many external users" weight */
    MANY: 5,
};
